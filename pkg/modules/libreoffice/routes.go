package libreoffice

import (
	"errors"
	"fmt"
	"io/fs"
	"net/http"
	"os"
	"path/filepath"

	"github.com/google/uuid"
	"github.com/gotenberg/gotenberg/v7/pkg/gotenberg"
	"github.com/gotenberg/gotenberg/v7/pkg/modules/api"
	"github.com/gotenberg/gotenberg/v7/pkg/modules/libreoffice/uno"
	"github.com/labstack/echo/v4"
)

// convertRoute returns an api.Route which can convert LibreOffice documents
// to PDF.
func convertRoute(unoAPI uno.API, engine gotenberg.PDFEngine) api.Route {
	return api.Route{
		Method:      http.MethodPost,
		Path:        "/forms/libreoffice/convert",
		IsMultipart: true,
		Handler: func(c echo.Context) error {
			ctx := c.Get("context").(*api.Context)

			// Let's get the data from the form and validate them.
			var (
				inputPaths         []string
				landscape          bool
				nativePageRanges   string
				nativePDFA1aFormat bool
				nativePDFformat    string
				PDFformat          string
				merge              bool
				asImages           bool
				slideImageDensity  string
				slideImageQuality  string
				slideImageResize   string
			)

			err := ctx.FormData().
				MandatoryPaths(unoAPI.Extensions(), &inputPaths).
				Bool("landscape", &landscape, false).
				String("nativePageRanges", &nativePageRanges, "").
				Bool("nativePdfA1aFormat", &nativePDFA1aFormat, false).
				String("nativePdfFormat", &nativePDFformat, "").
				String("pdfFormat", &PDFformat, "").
				Bool("merge", &merge, false).
				Bool("asImages", &asImages, false).
				// These defaults seem to produce a reasonably good quality
				String("slideImageDensity", &slideImageDensity, "288").
				String("slideImageQuality", &slideImageQuality, "85").
				// Rendering at a higher density and then reducing size seems to produce better quality
				String("slideImageResize", &slideImageResize, "50%").
				Validate()

			if err != nil {
				return fmt.Errorf("validate form data: %w", err)
			}

			if nativePDFA1aFormat {
				ctx.Log().Warn("'nativePdfA1aFormat' is deprecated; prefer 'nativePdfFormat' or 'pdfFormat' form fields instead")
			}

			if nativePDFA1aFormat && nativePDFformat != "" {
				return api.WrapError(
					errors.New("got both 'nativePdfFormat' and 'nativePdfA1aFormat' form fields"),
					api.NewSentinelHTTPError(http.StatusBadRequest, "Both 'nativePdfFormat' and 'nativePdfA1aFormat' form fields are provided"),
				)
			}

			if nativePDFA1aFormat && PDFformat != "" {
				return api.WrapError(
					errors.New("got both 'pdfFormat' and 'nativePdfA1aFormat' form fields"),
					api.NewSentinelHTTPError(http.StatusBadRequest, "Both 'pdfFormat' and 'nativePdfA1aFormat' form fields are provided"),
				)
			}

			if nativePDFformat != "" && PDFformat != "" {
				return api.WrapError(
					errors.New("got both 'pdfFormat' and 'nativePdfFormat' form fields"),
					api.NewSentinelHTTPError(http.StatusBadRequest, "Both 'pdfFormat' and 'nativePdfFormat' form fields are provided"),
				)
			}

			if asImages && len(inputPaths) > 1 {
				return api.WrapError(
					errors.New("multiple input files are not supported when converting to images"),
					api.NewSentinelHTTPError(http.StatusBadRequest, fmt.Sprintf("there should be only one input file when converting to images")),
				)
			}

			if nativePDFA1aFormat {
				nativePDFformat = gotenberg.FormatPDFA1a
			}

			// Alright, let's convert each document to PDF.

			outputPaths := make([]string, len(inputPaths))

			ctx.Log().Info("Converting input to PDF...")
			for i, inputPath := range inputPaths {
				outputPaths[i] = ctx.GeneratePath(".pdf")

				options := uno.Options{
					Landscape:  landscape,
					PageRanges: nativePageRanges,
					PDFformat:  nativePDFformat,
				}

				err = unoAPI.PDF(ctx, ctx.Log(), inputPath, outputPaths[i], options)

				if err != nil {
					if errors.Is(err, uno.ErrMalformedPageRanges) {
						return api.WrapError(
							fmt.Errorf("convert to PDF: %w", err),
							api.NewSentinelHTTPError(http.StatusBadRequest, fmt.Sprintf("Malformed page ranges '%s' (nativePageRanges)", options.PageRanges)),
						)
					}

					return fmt.Errorf("convert to PDF: %w", err)
				}
			}
			ctx.Log().Info("Finished converting to PDF")

			// So far so good, let's check if we have to merge the PDFs. Quick
			// win: if there is only one PDF, skip this step.

			if len(outputPaths) > 1 && merge {
				outputPath := ctx.GeneratePath(".pdf")

				err = engine.Merge(ctx, ctx.Log(), outputPaths, outputPath)
				if err != nil {
					return fmt.Errorf("merge PDFs: %w", err)
				}

				// Now, let's check if the client want to convert this result
				// PDF to a specific PDF format.

				// Note: nativePdfA1aFormat/nativePdfFormat have not been
				// specified if PDFformat is not empty.

				if PDFformat != "" {
					convertInputPath := outputPath
					convertOutputPath := ctx.GeneratePath(".pdf")

					err = engine.Convert(ctx, ctx.Log(), PDFformat, convertInputPath, convertOutputPath)

					if err != nil {
						if errors.Is(err, gotenberg.ErrPDFFormatNotAvailable) {
							return api.WrapError(
								fmt.Errorf("convert PDF: %w", err),
								api.NewSentinelHTTPError(
									http.StatusBadRequest,
									fmt.Sprintf("At least one PDF engine does not handle the PDF format '%s' (pdfFormat), while other have failed to convert for other reasons", PDFformat),
								),
							)
						}

						return fmt.Errorf("convert PDF: %w", err)
					}

					// Important: the output path is now the converted file.
					outputPath = convertOutputPath
				}

				// Last but not least, add the output path to the context so that
				// the API is able to send it as a response to the client.

				err = ctx.AddOutputPaths(outputPath)
				if err != nil {
					return fmt.Errorf("add output path: %w", err)
				}

				return nil
			}

			// Ok, we don't have to merge the PDFs. Let's check if the client
			// want to convert each PDF to a specific PDF format.

			// Note: nativePdfA1aFormat/nativePdfFormat have not been
			// specified if PDFformat is not empty.

			if PDFformat != "" {
				convertOutputPaths := make([]string, len(outputPaths))

				for i, outputPath := range outputPaths {
					convertInputPath := outputPath
					convertOutputPaths[i] = ctx.GeneratePath(".pdf")

					err = engine.Convert(ctx, ctx.Log(), PDFformat, convertInputPath, convertOutputPaths[i])

					if err != nil {
						if errors.Is(err, gotenberg.ErrPDFFormatNotAvailable) {
							return api.WrapError(
								fmt.Errorf("convert PDF: %w", err),
								api.NewSentinelHTTPError(
									http.StatusBadRequest,
									fmt.Sprintf("At least one PDF engine does not handle the PDF format '%s' (pdfFormat), while other have failed to convert for other reasons", PDFformat),
								),
							)
						}

						return fmt.Errorf("convert PDF: %w", err)
					}

				}

				// Important: the output paths are now the converted files.
				outputPaths = convertOutputPaths
			}

			if asImages {
				resultDir := filepath.Join(filepath.Dir(outputPaths[0]), uuid.NewString())
				err := os.MkdirAll(resultDir, 0755)
				if err != nil {
					return fmt.Errorf("cannot create result folder: %w", err)
				}

				outputFilePath := filepath.Join(resultDir, "slide.jpg")

				args := []string{
					"-density",
					slideImageDensity,
					outputPaths[0],
					"-quality",
					slideImageQuality,
					"-resize",
					slideImageResize,
					outputFilePath,
				}

				ctx.Log().Info("Creating slide images out of the resulting PDF...")
				convertCmd, err := gotenberg.CommandContext(ctx, ctx.Log(), "/usr/bin/convert", args...)
				if err != nil {
					return api.WrapError(
						fmt.Errorf("failed to build a command for conversion to images: %w", err),
						api.NewSentinelHTTPError(http.StatusBadRequest, fmt.Sprintf("failed to build a command for conversion to images")),
					)
				}

				// Uncomment this block if there is a need to inspect command output
				//convertCmd := exec.CommandContext(ctx, "/usr/bin/convert", args...)
				//var outBuffer, errBuffer bytes.Buffer
				//convertCmd.Stdout = &outBuffer
				//convertCmd.Stderr = &errBuffer

				//err = convertCmd.Run()
				//if err != nil {
				//	ctx.Log().Error("> > > COMMAND WAS: " + convertCmd.String())
				//	ctx.Log().Error("> > > STDOUT: " + outBuffer.String())
				//	ctx.Log().Error("> > > STD ERR: " + errBuffer.String())
				//	return fmt.Errorf("failed to convert pdf to images: %w", err)
				//}

				exitCode, err := convertCmd.Exec()

				if err != nil {
					ctx.Log().Error("> > COMMAND WAS: " + convertCmd.CmdString())
					return fmt.Errorf("failed to create images from PDF: %w, exit code: %d", err, exitCode)
				}
				ctx.Log().Info("Done creating images")

				var resultPaths []string

				err = filepath.WalkDir(resultDir, func(path string, info fs.DirEntry, err error) error {
					if err != nil {
						return err
					}
					if info.IsDir() {
						// Skip folders, need images only
						return nil
					}

					resultPaths = append(resultPaths, path)
					return nil
				})

				if err != nil {
					return fmt.Errorf("failed to return created images: %w", err)
				}

				ctx.Log().Info("Writing JSON data...")
				dataCmd, err := gotenberg.CommandContext(
					ctx,
					ctx.Log(),
					"/usr/bin/python",
					"/usr/bin/write_slide_data.py",
					inputPaths[0],
					resultDir,
				)
				if err != nil {
					return fmt.Errorf("failed to create a command that writes slide data: %w", err)
				}

				//dataCmd := exec.CommandContext(ctx, "/usr/bin/python", "/usr/bin/write_slide_data.py", inputPaths[0], resultDir)
				//var pyOut, pyErr bytes.Buffer
				//dataCmd.Stdout = &pyOut
				//dataCmd.Stderr = &pyErr

				_, err = dataCmd.Exec()
				if err != nil {
					//ctx.Log().Error("> > > PYTHON SCRIPT FAILED ")
					//ctx.Log().Error("> > > OUTPUT: " + pyOut.String())
					//ctx.Log().Error("> > > ERROR: " + pyErr.String())
					return fmt.Errorf("failed to write slide data: %w", err)
				}
				resultPaths = append(resultPaths, filepath.Join(resultDir, "data.json"))
				ctx.Log().Info("Done writing JSON data")

				err = ctx.AddOutputPaths(resultPaths...)
			} else {
				// Last but not least, add the output paths to the context so that
				// the API is able to send them as a response to the client.
				err = ctx.AddOutputPaths(outputPaths...)
			}

			if err != nil {
				return fmt.Errorf("add output paths: %w", err)
			}

			return nil
		},
	}
}
