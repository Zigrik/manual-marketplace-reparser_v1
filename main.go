package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"sort"
	"strings"

	"github.com/xuri/excelize/v2"
)

type Product struct {
	Name       string
	Code       string
	Stock      string
	PriceShina string
	PriceMin   string
	SourceFile string
	Link       string
}

func main() {
	// Создаем выходной файл
	outputFile := excelize.NewFile()
	outputSheet := "Объединенные данные"
	index, _ := outputFile.NewSheet(outputSheet)
	outputFile.SetActiveSheet(index)
	outputFile.DeleteSheet("Sheet1")

	// Заголовки
	headers := []string{"Номенклатура", "Код", "Остаток", "цена Шинорама", "цена минимум", "Источник", "ссылка"}
	for col, header := range headers {
		cell, _ := excelize.CoordinatesToCellName(col+1, 1)
		outputFile.SetCellValue(outputSheet, cell, header)
	}

	var products []Product
	filesProcessed := 0
	foundTDSheet := false

	// Чтение файлов
	err := filepath.Walk("files", func(path string, info os.FileInfo, err error) error {
		if err != nil {
			log.Printf("Ошибка доступа к пути %s: %v", path, err)
			return nil
		}

		if info.IsDir() || !strings.HasSuffix(strings.ToLower(path), ".xlsx") {
			return nil
		}

		f, err := excelize.OpenFile(path)
		if err != nil {
			log.Printf("Ошибка открытия файла %s: %v", path, err)
			return nil
		}
		defer f.Close()

		// Проверяем наличие листа TDSheet
		sheetExists := false
		for _, sheet := range f.GetSheetList() {
			if sheet == "TDSheet" {
				sheetExists = true
				break
			}
		}

		if !sheetExists {
			log.Printf("Файл %s не содержит листа TDSheet", path)
			return nil
		}

		foundTDSheet = true
		rows, err := f.GetRows("TDSheet")
		if err != nil {
			log.Printf("Ошибка чтения листа TDSheet из %s: %v", path, err)
			return nil
		}

		if len(rows) <= 1 {
			log.Printf("Лист TDSheet в файле %s пуст или содержит только заголовки", path)
			return nil
		}

		// Более гибкая проверка заголовков
		expectedHeaders := map[string]bool{
			"номенклатура": false,
			"код":          false,
			"остаток":      false,
			"шинорама":     false,
			"минимум":      false,
			"ссылка":       false,
		}

		// Проверяем заголовки (первая строка)
		for _, h := range rows[0] {
			hLower := strings.ToLower(h)
			for expected := range expectedHeaders {
				if strings.Contains(hLower, expected) {
					expectedHeaders[expected] = true
				}
			}
		}

		// Проверяем что все нужные заголовки найдены
		for expected, found := range expectedHeaders {
			if !found {
				log.Printf("В файле %s не найден заголовок содержащий '%s'", path, expected)
				return nil
			}
		}

		// Определяем индексы столбцов
		colIndexes := make(map[string]int)
		for i, h := range rows[0] {
			hLower := strings.ToLower(h)
			switch {
			case strings.Contains(hLower, "номенклатура"):
				colIndexes["name"] = i
			case strings.Contains(hLower, "код"):
				colIndexes["code"] = i
			case strings.Contains(hLower, "остаток"):
				colIndexes["stock"] = i
			case strings.Contains(hLower, "шинорама"):
				colIndexes["price_shina"] = i
			case strings.Contains(hLower, "минимум"):
				colIndexes["price_min"] = i
			case strings.Contains(hLower, "ссылка"):
				colIndexes["link"] = i
			}
		}

		// Обрабатываем строки данных
		for rowIdx, row := range rows[1:] {
			// Проверяем что строка содержит достаточно данных
			requiredCols := []string{"name", "code", "stock", "price_shina", "price_min", "link"}
			valid := true
			for _, col := range requiredCols {
				if colIndexes[col] >= len(row) {
					log.Printf("Пропуск строки %d в файле %s: недостаточно столбцов", rowIdx+2, path)
					valid = false
					break
				}
			}
			if !valid {
				continue
			}

			products = append(products, Product{
				Name:       strings.TrimSpace(row[colIndexes["name"]]),
				Code:       strings.TrimSpace(row[colIndexes["code"]]),
				Stock:      strings.TrimSpace(row[colIndexes["stock"]]),
				PriceShina: strings.TrimSpace(row[colIndexes["price_shina"]]),
				PriceMin:   strings.TrimSpace(row[colIndexes["price_min"]]),
				Link:       strings.TrimSpace(row[colIndexes["link"]]),
				SourceFile: filepath.Base(path),
			})
		}

		filesProcessed++
		log.Printf("Обработан файл: %s (%d строк)", path, len(rows)-1)
		return nil
	})

	if err != nil {
		log.Fatal("Ошибка при обходе папки:", err)
	}

	if !foundTDSheet {
		log.Fatal("Ни один файл не содержит листа с именем TDSheet")
	}

	if filesProcessed == 0 {
		log.Fatal("Не найдено ни одного подходящего файла в папке 'files'")
	}

	if len(products) == 0 {
		log.Fatal("Не найдено ни одной строки данных во всех файлах")
	}

	// Сортировка
	sort.Slice(products, func(i, j int) bool {
		return strings.ToLower(products[i].Name) < strings.ToLower(products[j].Name)
	})

	// Запись данных
	for rowIdx, product := range products {
		data := []string{
			product.Name,
			product.Code,
			product.Stock,
			product.PriceShina,
			product.PriceMin,
			product.Link,
			product.SourceFile,
		}

		for colIdx, value := range data {
			cell, _ := excelize.CoordinatesToCellName(colIdx+1, rowIdx+2)
			if err := outputFile.SetCellValue(outputSheet, cell, value); err != nil {
				log.Printf("Ошибка записи в ячейку %s: %v", cell, err)
			}
		}
	}

	// Автоширина столбцов
	for col := 1; col <= len(headers); col++ {
		colName, _ := excelize.ColumnNumberToName(col)
		outputFile.SetColWidth(outputSheet, colName, colName, 20)
	}

	if err := outputFile.SaveAs("объединенные_данные.xlsx"); err != nil {
		log.Fatal("Ошибка сохранения файла:", err)
	}

	fmt.Printf("Готово! Обработано %d файлов, %d строк данных\n", filesProcessed, len(products))
	fmt.Println("Результат сохранен в 'объединенные_данные.xlsx'")
}
