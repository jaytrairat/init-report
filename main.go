package main

import (
	"fmt"
	"os"
	"time"

	"github.com/nguyenthenguyen/docx"
	"github.com/spf13/cobra"
)

var (
	template     string
	caseNumber   string
	evidenceList []string
)

var rootCmd = &cobra.Command{
	Use:   "init-report",
	Short: "Generate DFU Report",
	Run: func(cmd *cobra.Command, args []string) {
		fmt.Println(caseNumber)

		r, err := docx.ReadDocxFile(template)
		if err != nil {
			fmt.Println("ERROR :: Cannot create report")
			os.Exit(0)
		}
		result := r.Editable()
		result.Replace("valCaseNumber", caseNumber, -1)

		timestamp := time.Now().Format("2006-01-02_15-04-05")

		outputFileName := fmt.Sprintf("%s.docx", timestamp)
		outputFilePath := fmt.Sprintf("%s", outputFileName)

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
	rootCmd.Flags().StringVarP(&template, "template", "t", "ReportTemplate.docx", "Report template")
	rootCmd.Flags().StringVarP(&caseNumber, "casenumber", "c", "01/2566", "Case number (required)")
	rootCmd.Flags().StringSliceVarP(&evidenceList, "evidence", "e", []string{}, "Evidence")
	rootCmd.MarkFlagRequired("casenumber")
	Execute()
}
