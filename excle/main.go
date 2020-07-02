package main

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"os"
	"time"
)

var (
	ModelNameXlsx = "model.xlsx"
	BeAccessList []BeAccess
	AccessList []string
)

type BeAccess struct {
	BeAccessName string
	Code string
}


func main() {
		readModel()
	update()
}


func update() {
	var file *xlsx.File
	fileName := time.Now().Format("2006-01-02")
	err:=os.Mkdir(fileName,0777)
	if err!=nil{
		fmt.Println(err)
		return
	}
	excelFileName := ModelNameXlsx
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
	for k, sheet := range xlFile.Sheets {
		if k == 0 {
			for _, vBeAccess := range BeAccessList {
				for _, vAccess := range AccessList {
					for kr, row := range sheet.Rows {
						for kc, cell := range row.Cells {
							if kr == 1 && kc == 2 {
								cell.Value = vBeAccess.BeAccessName //被评估人
							}
							if kr == 1 && kc == 4 {
								cell.Value = vBeAccess.Code  // 编码
							}
							if kr == 1 && kc == 7 {
								cell.Value = vAccess //被评估人
							}
							//写入新文件sheet
							file = xlsx.NewFile()
							newsheet := &sheet
							file.Sheets = append(file.Sheets, *newsheet)
							err = file.Save(fmt.Sprintf("%s/%s--%s.xlsx",fileName,vAccess,vBeAccess.BeAccessName))
							if err != nil {
								fmt.Printf(err.Error())
								return
							}
							fmt.Printf("%s--%s\n",vAccess,vBeAccess.BeAccessName)
						}
					}
				}
			}
		}
	}
	fmt.Println("success")
}

func readModel()  {
	excelFileName := ModelNameXlsx
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		fmt.Printf("open failed: %s\n", err)
		return
	}
	for k, sheet := range xlFile.Sheets {
		if k == 1 { //第二页开始
			for kr, row := range sheet.Rows {
				if kr == 0 {
					continue
				}
				for kc, cell := range row.Cells {
					if kc == 0 {
						if cell.Value == "" {
							continue
						}
						AccessList = append(AccessList, cell.Value)
					}
					if kc == 1 {
						BeAccessList = append(BeAccessList, BeAccess{
							BeAccessName: cell.Value,
							Code:         row.Cells[kc+1].Value,
						})
					}
					//fmt.Printf("[%d]--[%d]-->%s\n",kr, kc, cell.String())
				}
			}
		}
	}
	//fmt.Printf("评估人:%v", AccessList)
	//fmt.Printf("被评估人:%v", BeAccessList)
}


func read() {
	excelFileName := "aa.xlsx"
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		fmt.Printf("open failed: %s\n", err)
	}
	for _, sheet := range xlFile.Sheets {
		fmt.Printf("Sheet Name: %s\n", sheet.Name)
		for kr, row := range sheet.Rows {
			for kc, cell := range row.Cells {
				text := cell.String()
				fmt.Printf("行列[%d]->[%d]",kr,kc)
				fmt.Printf("%s\n", text)
			}
		}
	}
}





