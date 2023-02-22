package main

import (
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/tealeg/xlsx"
)

func pegarUltimoArquivoExcel() string {
	// Lista todos os arquivos na pasta atual
	files, err := filepath.Glob("*")
	if err != nil {
		fmt.Println("Erro ao listar arquivos:", err)
		return ""
	}

	// Filtra apenas os arquivos xlsx
	var xlsxFiles []string
	for _, file := range files {
		if strings.HasSuffix(file, ".xlsx") {
			xlsxFiles = append(xlsxFiles, file)
		}
	}

	// Verifica se tem arquivo xlsx na lista
	if len(xlsxFiles) > 0 {
		// Ordena a lista de arquivos por data de modificação (mais recente primeiro)
		sort.Slice(xlsxFiles, func(i, j int) bool {
			info1, _ := os.Stat(xlsxFiles[i])
			info2, _ := os.Stat(xlsxFiles[j])
			return info1.ModTime().After(info2.ModTime())
		})

		// Pega o nome do primeiro arquivo xlsx na lista (mais recente)
		var lastXlsxFile string
		if len(xlsxFiles) > 0 {
			lastXlsxFile = xlsxFiles[0]
		}

		return lastXlsxFile
	} else {
		return ""
	}

}

func main() {
	arquivoXlsx := pegarUltimoArquivoExcel()

	// Verifica se encontrou algum arquivo xlsx
	if arquivoXlsx != "" {
		// abre o arquivo xlsx
		xlFile, err := xlsx.OpenFile(arquivoXlsx)
		if err != nil {
			fmt.Println("Erro ao abrir arquivo:", err)
			return
		}

		// pega a primeira aba
		sheet := xlFile.Sheets[0]

		// Pega o nome da primeira aba
		sheetName := xlFile.Sheets[0].Name

		// realiza a somatória dos valores da coluna F a partir da 10 linha
		sum := 0.0
		lastRow := 0
		for i, row := range sheet.Rows {
			if i < 10 {
				continue // ignora as primeiras 10 linhas
			}
			cell := row.Cells[5] // coluna F (a primeira coluna tem índice 0)
			if cell != nil {
				// remove a letra "h" da célula (se existir) e converte para número
				valueStr := strings.ReplaceAll(cell.String(), " h", "")
				value, _ := strconv.ParseFloat(valueStr, 64)
				sum += value
				lastRow = i + 1
			}
		}

		f, err := excelize.OpenFile(arquivoXlsx)

		if err != nil {
			fmt.Println("Erro ao abrir arquivo para escrita:", err)
			return
		}

		// adiciona a somatória após o último valor da coluna F
		f.SetCellValue(sheetName, fmt.Sprintf("F%d", lastRow+1), sum)

		for i := 11; i < lastRow+1; i++ {
			celula := f.GetCellValue(sheetName, fmt.Sprintf("F%d", i))
			valueStr := strings.ReplaceAll(celula, " h", "")
			value, _ := strconv.ParseFloat(valueStr, 64)
			f.SetCellValue(sheetName, fmt.Sprintf("F%d", i), value)
		}

		// salva as alterações no arquivo
		err = f.Save()
		if err != nil {
			fmt.Println("Erro ao salvar alterações:", err)
		}
	} else {
		fmt.Println("Não foi encontrado nenhum arquivo com extensão .xlsx")
	}
}
