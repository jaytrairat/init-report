package main

import (
	"fmt"
	"os"
	"strings"
	"time"

	"github.com/nguyenthenguyen/docx"
	"github.com/spf13/cobra"
)

var (
	template     string
	caseNumber   string
	evidenceList []string
	issueList    []string
)

var rootCmd = &cobra.Command{
	Use:   "init-report",
	Short: "Generate DFU Report",
	Run: func(cmd *cobra.Command, args []string) {
		r, err := docx.ReadDocxFile(template)
		if err != nil {
			fmt.Println("ERROR :: Cannot create report")
			os.Exit(0)
		}
		result := r.Editable()

		var evidenceInString strings.Builder
		evidenceIndex := 1
		for _, evidence := range evidenceList {
			evidenceInString.WriteString(fmt.Sprintf("พยานหลักฐานลำดับที่ %d %s ยี่ห้อ xxx รุ่น xxx", evidenceIndex, evidence))
			evidenceIndex++
		}

		var issueInString strings.Builder
		issueIndex := 1
		for _, issue := range issueList {
			issueInString.WriteString(fmt.Sprintf("2.%d %s", issueIndex, issue))
			issueIndex++
		}

		result.Replace("valCaseNumber", caseNumber, -1)
		result.Replace("valListOfEvidence", evidenceInString.String(), -1)
		result.Replace("valListOfIssue", issueInString.String(), -1)

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
	rootCmd.Flags().StringSliceVarP(&issueList, "issue", "i", []string{}, "Issue")
	Execute()
}
