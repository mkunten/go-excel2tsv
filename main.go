package main

import (
	"errors"
	"fmt"
	"github.com/comail/colog"
	"github.com/deiwin/interact"
	"github.com/tealeg/xlsx"
	"io/ioutil"
	"log"
	"os"
	"path/filepath"
	"strings"
)

var flagRewriteAll bool

func usage() {
	fmt.Printf("Usage: %s [files]", os.Args[0])
	os.Exit(1)
}

func checkPrompt(input string) error {
	if strings.Index("ysacYSAC", input) == -1 {
		return errors.New("select from [y/s/a/c]")
	}
	return nil
}

func main() {
	// set colog
	colog.SetDefaultLevel(colog.LInfo)
	colog.SetMinLevel(colog.LInfo)
	//colog.SetMinLevel(colog.LTrace)
	colog.SetFormatter(&colog.StdFormatter{
		Colors: true,
		Flag:   log.Ldate | log.Ltime | log.Lshortfile,
	})
	colog.Register()
	//
	actor := interact.NewActor(os.Stdin, os.Stdout)

	if len(os.Args) == 1 {
		usage()
	}
	args := os.Args[1:]
	for _, excelFileName := range args {
		if filepath.Ext(excelFileName) != ".xlsx" {
			log.Printf("warn: not an excel file: %s\n", excelFileName)
			continue
		}
		baseName := excelFileName[0 : len(excelFileName)-5]
		tsvFileName := ""

		xlFile, err := xlsx.OpenFile(excelFileName)
		if err != nil {
			log.Printf("error: %v\n", err)
		}
		for sIdx, sheet := range xlFile.Sheets {
			//
			if sIdx == 0 && len(xlFile.Sheets) == 1 {
				tsvFileName = baseName + ".tsv"
			} else {
				tsvFileName = baseName + "-" + sheet.Name + ".tsv"
			}
			if !flagRewriteAll {
				_, err := os.Stat(tsvFileName)
				if err == nil {
					r, err := actor.PromptOptional(fmt.Sprintf("%s already exists, rewrite? [y(es)/S(kip)/a(ll)/c(ancel)]", tsvFileName), "s", checkPrompt)
					for err != nil {
						fmt.Printf("Please select from y/s/a/c\n")
						r, err = actor.PromptOptional(fmt.Sprintf("%s already exists, rewrite? [y(es)/S(kip)/a(ll)/c(ancel)]", tsvFileName), "s", checkPrompt)
					}
					switch r {
					case "a", "A":
						flagRewriteAll = true
					case "s", "S":
						continue
					case "c", "C":
						log.Printf("canceled.\n")
						os.Exit(0)
					}
				}
			}
			//
			lines := make([]string, len(sheet.Rows))
			for rIdx, row := range sheet.Rows {
				line := make([]string, len(row.Cells))
				for cIdx, cell := range row.Cells {
					line[cIdx], _ = cell.String()
				}
				lines[rIdx] = strings.Join(line, "\t")
			}
			err = ioutil.WriteFile(tsvFileName,
				[]byte(strings.Join(lines, "\n")), 0644)
			if err != nil {
				log.Printf("error: %v\n", err)
			}
			log.Printf("%s created", tsvFileName)
		}
	}
	fmt.Print("done\n")
}
