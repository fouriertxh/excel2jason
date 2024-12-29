package main

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/storage"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
)

func main() {
	myApp := app.New()
	myWindow := myApp.NewWindow("Excel 和 JSON 转换工具")
	myWindow.Resize(fyne.NewSize(500, 300))

	importPathLabel := widget.NewLabel("导入路径：未选择")
	exportPathLabel := widget.NewLabel("导出路径：未选择")

	// 导入 Excel 转 JSON 按钮
	importButton := widget.NewButton("导入 Excel 转 JSON", func() {
		homeDir, _ := os.UserHomeDir()
		defaultDirURI := storage.NewFileURI(homeDir)

		// 设置为 ListableURI
		defaultDir, err := storage.ListerForURI(defaultDirURI)
		if err != nil {
			dialog.ShowError(fmt.Errorf("无法设置默认目录: %v", err), myWindow)
			return
		}

		// 文件打开对话框
		openDialog := dialog.NewFileOpen(func(reader fyne.URIReadCloser, err error) {
			if err != nil || reader == nil {
				return
			}

			inputPath := reader.URI().Path()
			importPathLabel.SetText(fmt.Sprintf("导入路径：%s", inputPath))

			// 设置默认导出路径为与输入路径相同的目录
			outputPath := filepath.Join(filepath.Dir(inputPath), "output.json")
			err = excelToJSON(inputPath, outputPath)
			if err != nil {
				dialog.ShowError(err, myWindow)
				return
			}
			dialog.ShowInformation("成功", fmt.Sprintf("Excel 转 JSON 成功！保存至：%s", outputPath), myWindow)
		}, myWindow)

		openDialog.SetFilter(storage.NewExtensionFileFilter([]string{".xlsx"}))
		openDialog.SetLocation(defaultDir)
		openDialog.Show()
	})

	// 导出 JSON 转 Excel 按钮
	exportButton := widget.NewButton("导出 JSON 转 Excel", func() {
		openDialog := dialog.NewFileOpen(func(reader fyne.URIReadCloser, err error) {
			if err != nil || reader == nil {
				return
			}

			inputPath := reader.URI().Path()
			if filepath.Ext(inputPath) != ".json" {
				dialog.ShowError(fmt.Errorf("请选择有效的 JSON 文件"), myWindow)
				return
			}
			exportPathLabel.SetText(fmt.Sprintf("输入路径：%s", inputPath))

			saveDialog := dialog.NewFileSave(func(writer fyne.URIWriteCloser, err error) {
				if err != nil || writer == nil {
					return
				}

				outputPath := writer.URI().Path()
				if filepath.Ext(outputPath) != ".xlsx" {
					outputPath += ".xlsx"
				}

				exportPathLabel.SetText(fmt.Sprintf("导出路径：%s", outputPath))

				err = jsonToExcel(inputPath, outputPath)
				if err != nil {
					dialog.ShowError(err, myWindow)
					return
				}
				dialog.ShowInformation("成功", fmt.Sprintf("JSON 转 Excel 成功！保存至：%s", outputPath), myWindow)
			}, myWindow)

			saveDialog.SetFilter(storage.NewExtensionFileFilter([]string{".xlsx"}))
			saveDialog.Show()
		}, myWindow)

		openDialog.SetFilter(storage.NewExtensionFileFilter([]string{".json"}))
		openDialog.Show()
	})

	content := container.NewVBox(
		importPathLabel,
		importButton,
		exportPathLabel,
		exportButton,
	)
	myWindow.SetContent(content)
	myWindow.ShowAndRun()
}

// 检查文件的写入权限
func checkWritePermission(path string) error {
	testFile := path + ".permission_test"
	file, err := os.Create(testFile)
	if err != nil {
		return fmt.Errorf("没有写入权限，无法创建文件: %v", err)
	}
	file.Close()
	err = os.Remove(testFile)
	if err != nil {
		return fmt.Errorf("无法删除测试文件，权限问题: %v", err)
	}
	return nil
}

// Excel 转 JSON
func excelToJSON(inputPath, outputPath string) error {
	err := checkWritePermission(outputPath)
	if err != nil {
		return err
	}

	f, err := excelize.OpenFile(inputPath)
	if err != nil {
		return fmt.Errorf("打开 Excel 文件失败: %v", err)
	}
	defer f.Close()

	sheetNames := f.GetSheetList()
	if len(sheetNames) == 0 {
		return fmt.Errorf("Excel 文件中没有有效的工作表")
	}

	sheetName := sheetNames[0]
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("读取行数据失败: %v", err)
	}

	if len(rows) < 2 {
		return fmt.Errorf("工作表 '%s' 中没有有效的数据", sheetName)
	}

	// 将数据转换为 JSON 数组
	var data []map[string]string
	headers := rows[0] // 第一行作为表头
	for _, row := range rows[1:] {
		rowData := make(map[string]string)
		for i, cell := range row {
			if i < len(headers) {
				rowData[headers[i]] = cell
			}
		}
		data = append(data, rowData)
	}

	// 将 JSON 数组写入文件
	jsonData, err := json.MarshalIndent(data, "", "  ")
	if err != nil {
		return fmt.Errorf("JSON 编码失败: %v", err)
	}
	err = os.WriteFile(outputPath, jsonData, 0644)
	if err != nil {
		return fmt.Errorf("保存 JSON 文件失败: %v", err)
	}
	return nil
}

// JSON 转 Excel
func jsonToExcel(inputPath, outputPath string) error {
	jsonData, err := os.ReadFile(inputPath)
	if err != nil {
		return fmt.Errorf("读取 JSON 文件失败: %v", err)
	}

	var data []map[string]string
	err = json.Unmarshal(jsonData, &data)
	if err != nil {
		return fmt.Errorf("解析 JSON 数据失败: %v", err)
	}

	f := excelize.NewFile()
	defer f.Close()

	if len(data) > 0 {
		// 提取表头
		headers := []string{}
		for key := range data[0] {
			headers = append(headers, key)
		}

		// 写入表头到第一行
		for i, header := range headers {
			colName, _ := excelize.ColumnNumberToName(i + 1)
			f.SetCellValue("Sheet1", fmt.Sprintf("%s1", colName), header)
		}

		// 写入数据行
		for rowIdx, row := range data {
			for colIdx, header := range headers {
				colName, _ := excelize.ColumnNumberToName(colIdx + 1)
				f.SetCellValue("Sheet1", fmt.Sprintf("%s%d", colName, rowIdx+2), row[header])
			}
		}
	}

	err = f.SaveAs(outputPath)
	if err != nil {
		return fmt.Errorf("保存 Excel 文件失败: %v", err)
	}
	return nil
}
