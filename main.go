package main

import (
	"bufio"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"sort"
	"strings"

	"github.com/xuri/excelize/v2"
)

type Product struct {
	Data       map[string]string // Хранит все данные для каждой колонки
	SourceFile string            // Название файла-источника
}

func main() {
	// Читаем список дополнительных колонок из файла column.txt
	additionalColumns, err := readColumnsFromFile("column.txt")
	if err != nil {
		log.Fatalf("Ошибка чтения файла column.txt: %v", err)
	}

	// Создаем выходной файл
	outputFile := excelize.NewFile()
	outputSheet := "Объединенные данные"
	index, _ := outputFile.NewSheet(outputSheet)
	outputFile.SetActiveSheet(index)
	outputFile.DeleteSheet("Sheet1")

	// Базовые обязательные колонки + дополнительные из файла + источник
	headers := append([]string{"Наименование", "Код"}, additionalColumns...)
	headers = append(headers, "Источник")

	// Записываем заголовки
	for col, header := range headers {
		cell, _ := excelize.CoordinatesToCellName(col+1, 1)
		outputFile.SetCellValue(outputSheet, cell, header)
	}

	var products []Product
	filesProcessed := 0
	foundValidFiles := false

	// Чтение файлов
	err = filepath.Walk("files", func(path string, info os.FileInfo, err error) error {
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

		// Получаем список всех листов в файле
		sheets := f.GetSheetList()
		if len(sheets) == 0 {
			log.Printf("Файл %s не содержит ни одного листа", path)
			return nil
		}

		// Будем обрабатывать первый лист (можно изменить на цикл по всем листам)
		sheetName := sheets[0]
		rows, err := f.GetRows(sheetName)
		if err != nil {
			log.Printf("Ошибка чтения листа %s из %s: %v", sheetName, path, err)
			return nil
		}

		if len(rows) <= 1 {
			log.Printf("Лист %s в файле %s пуст или содержит только заголовки", sheetName, path)
			return nil
		}

		// Определяем индексы столбцов
		colIndexes := make(map[string]int)
		for i, h := range rows[0] {
			hLower := strings.ToLower(strings.TrimSpace(h))
			switch {
			case strings.Contains(hLower, "наименование"):
				colIndexes["Наименование"] = i
			case strings.Contains(hLower, "код"):
				colIndexes["Код"] = i
			default:
				// Проверяем дополнительные колонки
				for _, col := range additionalColumns {
					if strings.Contains(hLower, strings.ToLower(col)) {
						colIndexes[col] = i
						break
					}
				}
			}
		}

		// Проверяем наличие обязательных колонок
		if _, ok := colIndexes["Наименование"]; !ok {
			log.Printf("В файле %s не найден столбец 'Наименование'", path)
			return nil
		}
		if _, ok := colIndexes["Код"]; !ok {
			log.Printf("В файле %s не найден столбец 'Код'", path)
			return nil
		}

		foundValidFiles = true

		// Обрабатываем строки данных
		for _, row := range rows[1:] {
			product := Product{
				Data:       make(map[string]string),
				SourceFile: filepath.Base(path),
			}

			// Заполняем обязательные поля
			if colIndexes["Наименование"] < len(row) {
				product.Data["Наименование"] = strings.TrimSpace(row[colIndexes["Наименование"]])
			}
			if colIndexes["Код"] < len(row) {
				product.Data["Код"] = strings.TrimSpace(row[colIndexes["Код"]])
			}

			// Заполняем дополнительные поля
			for _, col := range additionalColumns {
				if colIdx, ok := colIndexes[col]; ok && colIdx < len(row) {
					product.Data[col] = strings.TrimSpace(row[colIdx])
				} else {
					product.Data[col] = "" // Пустая строка, если колонка не найдена
				}
			}

			products = append(products, product)
		}

		filesProcessed++
		log.Printf("Обработан файл: %s (%d строк)", path, len(rows)-1)
		return nil
	})

	if err != nil {
		log.Fatal("Ошибка при обходе папки:", err)
	}

	if !foundValidFiles {
		log.Fatal("Не найдено ни одного файла с обязательными колонками 'Наименование' и 'Код'")
	}

	if len(products) == 0 {
		log.Fatal("Не найдено ни одной строки данных во всех файлах")
	}

	// Сортировка по наименованию
	sort.Slice(products, func(i, j int) bool {
		return strings.ToLower(products[i].Data["Наименование"]) < strings.ToLower(products[j].Data["Наименование"])
	})

	// Запись данных
	for rowIdx, product := range products {
		for colIdx, header := range headers {
			value := ""
			if header == "Источник" {
				value = product.SourceFile
			} else {
				value = product.Data[header]
			}

			cell, _ := excelize.CoordinatesToCellName(colIdx+1, rowIdx+2)
			if err := outputFile.SetCellValue(outputSheet, cell, value); err != nil {
				log.Printf("Ошибка записи в ячейку %s: %v", cell, err)
			}
		}
	}

	// Устанавливаем ширину колонок
	for col := 1; col <= len(headers); col++ {
		colName, _ := excelize.ColumnNumberToName(col)
		width := 10
		header := headers[col-1]
		if header == "Наименование" || header == "Источник" || header == "комментарий" {
			width = 40
		}
		outputFile.SetColWidth(outputSheet, colName, colName, float64(width))
	}

	// Сохраняем файл
	if err := outputFile.SaveAs("объединенные_данные.xlsx"); err != nil {
		log.Fatal("Ошибка сохранения файла:", err)
	}

	fmt.Printf("\nГотово! Обработано %d файлов, %d строк данных\n", filesProcessed, len(products))
	fmt.Println("Результат сохранен в 'объединенные_данные.xlsx'")
	fmt.Println("\nНажмите любую клавишу для выхода...")
	waitForAnyKey()
}

// Функция для чтения списка колонок из файла
func readColumnsFromFile(filename string) ([]string, error) {
	file, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	var columns []string
	scanner := bufio.NewScanner(file)
	for scanner.Scan() {
		line := strings.TrimSpace(scanner.Text())
		if line != "" {
			columns = append(columns, line)
		}
	}

	if err := scanner.Err(); err != nil {
		return nil, err
	}

	return columns, nil
}

// Функция для ожидания нажатия любой клавиши
func waitForAnyKey() {
	bufio.NewReader(os.Stdin).ReadByte()
}
