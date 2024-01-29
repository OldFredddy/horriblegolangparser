package main

import (
	"encoding/xml"
	"fmt"
	"github.com/xuri/excelize/v2"
	"io/ioutil"
	"net/http"
	"os"
	"regexp"
	"strconv"
	"strings"
	"time"
)

func main() {
	urls, _ := getUrlsFromXML("ships.xml")
	year2 := urls[0]
	for i, url := range urls {
		if i != 0 {
			datesEventsMap, _ := fetchAndParseDates(url, year2)
			shipsName, _ := getShipName(url)
			saveToExcel(shipsName, "ships.xlsx", datesEventsMap, year2)
		}

	}
	url1 := "http://www.uscarriers.net/ddg55history.htm"
	year := "2023"
	removeEmptyRows("ships.xlsx")
	datesEventsMap, err := fetchAndParseDates(url1, year)
	if err != nil {
		fmt.Println("Ошибка:", err)
		return
	}

	for date, event := range datesEventsMap {
		fmt.Printf("[%s] [%s]\n", date, event)
	}
}
func removeHTMLTags(input string) string {
	re := regexp.MustCompile("<[^>]*>")
	return re.ReplaceAllString(input, "")
}

// Функция для парсинга дат и соответствующих событий на странице
func parseDatesAndEvents(htmlContent, year string) (map[string]string, error) {
	resultMap := make(map[string]string)
	months := []string{
		"January", "February", "March", "April",
		"May", "June", "July", "August",
		"September", "October", "November", "December",
	}

	nextYear, _ := strconv.Atoi(year)
	nextYearStr := strconv.Itoa(nextYear + 1)
	indexOfYear := strings.Index(htmlContent, year)
	indexOfNextYear := strings.Index(htmlContent, nextYearStr)

	if indexOfYear == -1 {
		return map[string]string{year: "year not found"}, nil
	}
	if indexOfNextYear == -1 {
		indexOfNextYear = len(htmlContent)
	}
	htmlContent = htmlContent[indexOfYear:indexOfNextYear]
	resultMap = extractDatesAndText(htmlContent, months)
	return resultMap, nil
}
func parseName(htmlContent string) (string, error) {
	re := regexp.MustCompile(`<span class="head[23]">(.*?)</span>`)
	matches := re.FindAllStringSubmatch(htmlContent, -1)
	var result string
	for _, match := range matches {
		result += match[1] + "\n"
	}
	return result, nil
}

// Функция, принимающая URL и год
func fetchAndParseDates(url, year string) (map[string]string, error) {
	resp, err := http.Get(url)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()

	body, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		return nil, err
	}

	return parseDatesAndEvents(string(body), year)
}
func getShipName(url string) (string, error) {
	resp, err := http.Get(url)
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()

	body, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		return "", err
	}

	return parseName(string(body))
}
func extractDatesAndText(htmlContent string, months []string) map[string]string {
	resultMap := make(map[string]string)
	htmlContent = deleteHTMLTags(htmlContent)
	monthsPattern := strings.Join(months, "|")
	re := regexp.MustCompile(`\b(` + monthsPattern + `)\s+(\d{1,2}),`)

	matches := re.FindAllStringSubmatchIndex(htmlContent, -1)
	if len(matches) == 0 {
		return resultMap // Возвращаем пустой resultMap, если совпадений нет
	}

	for i, match := range matches {
		date := htmlContent[match[0]:match[1]]

		// Определение endIndex: начало следующего совпадения или конец строки
		endIndex := len(htmlContent)
		if i < len(matches)-1 {
			nextMatch := matches[i+1]
			endIndex = nextMatch[0]
		}

		if match[1] > endIndex {
			continue // Пропускаем, если match[1] больше endIndex
		}

		text := htmlContent[match[1]:endIndex]
		resultMap[date] = deleteHTMLTags(text)
	}
	return resultMap
}
func deleteHTMLTags(input string) string {
	re := regexp.MustCompile("<[^>]*>")
	return re.ReplaceAllString(input, "")
}

// <strong>June 2</strong>
func deleteStrongTag(input string) string {
	re := regexp.MustCompile("<strong>*</strong>")
	return re.ReplaceAllString(input, "")
}

type URLSet struct {
	XMLName xml.Name `xml:"urls"`
	URLs    []URL    `xml:"url"`
}

type URL struct {
	Data string `xml:",chardata"`
}

func getUrlsFromXML(nameXML string) ([]string, error) {
	// Открытие файла
	file, err := os.Open(nameXML)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	// Чтение содержимого файла
	data, err := ioutil.ReadAll(file)
	if err != nil {
		return nil, err
	}

	// Декодирование XML
	var urlSet URLSet
	err = xml.Unmarshal(data, &urlSet)
	if err != nil {
		return nil, err
	}

	// Преобразование среза URL в срез строк
	urls := make([]string, len(urlSet.URLs))
	for i, url := range urlSet.URLs {
		urls[i] = url.Data
	}

	return urls, nil
}
func saveToExcel(shipName string, nameFile string, datesEventsMap map[string]string, yearStr string) {
	f, err := excelize.OpenFile(nameFile)
	if err != nil {
		f = excelize.NewFile()
	}

	year, _ := strconv.Atoi(yearStr)
	daysInYear := 365
	if year%4 == 0 && (year%100 != 0 || year%400 == 0) {
		daysInYear = 366
	}

	date := time.Date(year, time.January, 1, 0, 0, 0, 0, time.UTC)
	for day := 0; day < daysInYear; day++ {
		formattedDate := date.Format("02.01.2006")
		cell := fmt.Sprintf("A%d", day+1)
		f.SetCellValue("Sheet1", cell, formattedDate)
		date = date.AddDate(0, 0, 1)
	}

	colNum := 1
	for {
		cell := fmt.Sprintf("%c1", 'A'+colNum)
		if val, _ := f.GetCellValue("Sheet1", cell); val == "" {
			break
		}
		colNum++
	}

	shipNameCell := fmt.Sprintf("%c%d", 'A'+colNum, 1)
	f.SetCellValue("Sheet1", shipNameCell, shipName)

	for row := 1; row <= daysInYear; row++ {
		for date, val := range datesEventsMap {
			cell := fmt.Sprintf("A%d", row)
			cellDate, _ := f.GetCellValue("Sheet1", cell)
			if cellDate == convertDate(date, year) {
				eventCell := fmt.Sprintf("%c%d", 'A'+colNum, row)
				f.SetCellValue("Sheet1", eventCell, val)
				break
			}
		}
	}

	fillStyle, err := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#FFCCCC"},
			Pattern: 1,
		},
	})
	if err != nil {
		fmt.Println(err)
		return
	}

	// Проходим по всем ячейкам и применяем стиль к пустым ячейкам
	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}

	for rowNum, row := range rows {
		for colNum := range row {
			cell, _ := excelize.CoordinatesToCellName(colNum+1, rowNum+1)
			val, _ := f.GetCellValue("Sheet1", cell)
			if val == "" {
				f.SetCellStyle("Sheet1", cell, cell, fillStyle)
			}
		}
	}
	if err := f.SaveAs(nameFile); err != nil {
		fmt.Println(err)
	}
}

func convertDate(date string, year int) string {
	months := []string{
		"January", "February", "March", "April",
		"May", "June", "July", "August",
		"September", "October", "November", "December",
	}
	parts := strings.Split(strings.TrimRight(date, ","), " ")
	if len(parts) != 2 {
		return ""
	}
	monthStr, dayStr := parts[0], parts[1]
	var month int
	for idx, m := range months {
		if m == monthStr {
			month = idx + 1 // Месяцы в Go начинаются с 1
			break
		}
	}
	day, err := strconv.Atoi(dayStr)
	if err != nil {
		return "" // или обрабатывайте ошибку
	}

	// Формирование новой строки с датой
	newDate := time.Date(year, time.Month(month), day, 0, 0, 0, 0, time.UTC)
	return newDate.Format("02.01.2006")
}
func removeEmptyRows(nameFile string) {
	f, err := excelize.OpenFile(nameFile)
	if err != nil {
		fmt.Println("Ошибка при открытии файла:", err)
		return
	}

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println("Ошибка при получении строк:", err)
		return
	}

	for rowNum := len(rows) - 1; rowNum >= 0; rowNum-- {
		row := rows[rowNum]
		isEmpty := true
		for colNum, val := range row {
			if colNum == 0 {
				continue
			}
			if val != "" {
				isEmpty = false
				break
			}
		}

		if isEmpty {
			f.RemoveRow("Sheet1", rowNum+1)
		}
	}

	if err := f.SaveAs(nameFile); err != nil {
		fmt.Println("Ошибка при сохранении файла:", err)
	}
}
