package main

import (
	"fmt"
	"os"
	"time"

	"github.com/nguyenthenguyen/docx"
	"github.com/spf13/cobra"
)

var rootCmd = &cobra.Command{
	Use:   "init-report",
	Short: "Generate DFU Report",
	Run: func(cmd *cobra.Command, args []string) {
		caseNumber, _ := cmd.Flags().GetString("casenumber")

		r, err := docx.ReadDocxFile("./ReportTemplate.docx")
		if err != nil {
			panic(err)
		}
		result := r.Editable()
		result.Replace("valCaseNumber", caseNumber, -1)

		// Generate a timestamp in the "YYYY-MM-DD_HH-MM-SS" format.
		timestamp := time.Now().Format("2006-01-02_15-04-05")

		outputFilePath, _ := cmd.Flags().GetString("output")
		outputFileName := fmt.Sprintf("%s_%s.docx", caseNumber, timestamp)
		outputFilePath = fmt.Sprintf("%s/%s", outputFilePath, outputFileName)

		result.WriteToFile(outputFilePath)

		r.Close()

		fmt.Printf("Text replaced successfully. New DOCX file saved as %s\n", outputFilePath)
	},
}

func Execute() {
	if err := rootCmd.Execute(); err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
}

func main() {
	rootCmd.Flags().StringP("casenumber", "c", "", "Case number (required)")
	rootCmd.MarkFlagRequired("casenumber")
	Execute()
}
