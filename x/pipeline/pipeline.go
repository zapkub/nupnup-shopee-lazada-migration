package pipeline

import (
	"fmt"
	"io"

	"github.com/xuri/excelize/v2"
)

func New() *Pipeline {
	output := excelize.NewFile()
	return &Pipeline{
		output:                        output,
		imageURLMap:                   make(map[string][]string),
		descriptionMap:                make(map[string]string),
		priceMap:                      make(map[string]string),
		productColorOptionImageURLMap: make(map[string][]*ItemRowColor),
		excludeProductID:              make(map[string]struct{}),
	}
}

type Pipeline struct {
	output                        *excelize.File
	ExcludeProductIDSource        []*excelize.File
	SalesInfoSource               []*excelize.File
	MediaInfoSource               []*excelize.File
	BasicInfoSource               []*excelize.File
	PriceInfoSource               []*excelize.File
	imageURLMap                   map[string][]string
	productColorOptionImageURLMap map[string][]*ItemRowColor
	descriptionMap                map[string]string
	priceMap                      map[string]string
	excludeProductID              map[string]struct{}
}

func (p *Pipeline) SetCellValueToOutput(axis string, value interface{}) error {
	return p.output.SetCellValue("Sheet1", axis, value)
}

func (p *Pipeline) LoadBasicInfoFile(filenames []string) error {
	for _, name := range filenames {
		var file, err = excelize.OpenFile(name)
		if err != nil {
			return fmt.Errorf("cannot read basic info input from (%s) %w", name, err)
		}
		p.BasicInfoSource = append(p.BasicInfoSource, file)
	}
	return nil
}

func (p *Pipeline) LoadMediaInfoFile(filenames []string) error {
	for _, name := range filenames {
		var file, err = excelize.OpenFile(name)
		if err != nil {
			return fmt.Errorf("cannot read media info input from (%s) %w", name, err)
		}
		p.MediaInfoSource = append(p.MediaInfoSource, file)
	}
	return nil
}

func (p *Pipeline) LoadExcludeProductID(filenames []string) error {
	for _, name := range filenames {
		var file, err = excelize.OpenFile(name)
		if err != nil {
			return fmt.Errorf("cannot read exclude product id source input from (%s) %w", name, err)
		}
		p.ExcludeProductIDSource = append(p.ExcludeProductIDSource, file)
	}
	return nil
}

func (p *Pipeline) LoadSaleInfoFiles(filenames []string) error {
	for _, name := range filenames {
		var file, err = excelize.OpenFile(name)
		if err != nil {
			return fmt.Errorf("cannot read sale info input from (%s) %w", name, err)
		}
		p.SalesInfoSource = append(p.SalesInfoSource, file)
	}
	return nil
}

func (p *Pipeline) LoadPriceInfoFiles(filenames []string) error {
	for _, name := range filenames {
		var file, err = excelize.OpenFile(name)
		if err != nil {
			return fmt.Errorf("cannot read sale info input from (%s) %w", name, err)
		}
		p.PriceInfoSource = append(p.PriceInfoSource, file)
	}
	return nil
}

type ItemRowColor struct {
	ID       string
	ImageURL string
}
type OutputItemRow struct {
	ID          string
	Name        string
	Description string
	ImageURLs   []string
	Price       string
	Quantity    string
	Colors      []*ItemRowColor
}

func (p *Pipeline) extractExcludeIDFromSource() error {
	for _, excludeSource := range p.ExcludeProductIDSource {

		rows, err := excludeSource.Rows(excludeSource.GetSheetList()[0])
		if err != nil {
			return err
		}

		for rows.Next() {
			row, err := rows.Columns()
			if err != nil {
				return err
			}

			p.excludeProductID[row[0]] = struct{}{}
		}

	}
	return nil
}

var imageIndex = []int{4, 5, 6, 7, 8, 9, 10, 11, 12}
var imageColumns = []string{"D", "E", "F", "G", "H", "I", "J", "K", "L", "M"}

func (p *Pipeline) extractImageFromMediaSource() error {
	for _, media := range p.MediaInfoSource {
		rows, err := media.Rows("Sheet1")
		if err != nil {
			return err
		}

		for range [5]int{} {
			rows.Next()
			rows.Columns()
		}

		for rows.Next() {
			row, err := rows.Columns()
			if err != nil {
				return err
			}

			if len(row) == 0 {
				continue
			}

			var ID = row[0]
			p.imageURLMap[ID] = make([]string, 0, 10)
			for _, idx := range imageIndex {
				if idx > len(row)-1 {
					continue
				}
				p.imageURLMap[row[0]] = append(p.imageURLMap[row[0]], row[idx])
			}

			// iterate through the color variant columns
			// from P to BA column index from 15 to 60
			// each column will contain the value of
			// - column with n % 2 == 0 is imageURL
			// - column with n % 2 != 0 is color ID
			for i := 15; i < 60; i += 2 {

				if len(row)-1 < i {
					break
				}

				if row[i] == "" {
					break
				}

				var imageURL = ""
				if len(row) > i+1 {
					imageURL = row[i+1]
				}

				p.productColorOptionImageURLMap[ID] = append(
					p.productColorOptionImageURLMap[ID],
					&ItemRowColor{
						ID:       row[i],
						ImageURL: imageURL,
					},
				)
			}
		}
	}
	return nil
}

func (p *Pipeline) extractBasicInfoFromSource() error {
	for _, info := range p.BasicInfoSource {

		rows, err := info.Rows("Sheet1")
		if err != nil {
			return err
		}

		for rows.Next() {
			row, err := rows.Columns()
			if err != nil {
				return err
			}

			if len(row) == 0 {
				continue
			}
			if len(row) < 4 {
				continue
			}
			p.descriptionMap[row[0]] = row[3]

		}
	}
	return nil
}

func (p *Pipeline) extractPriceInfoFromSource() error {
	for _, info := range p.PriceInfoSource {

		rows, err := info.Rows("Sheet")
		if err != nil {
			return err
		}

		for rows.Next() {
			row, err := rows.Columns()
			if err != nil {
				return err
			}

			if len(row) == 0 {
				continue
			}
			if len(row) < 7 {
				continue
			}
			p.priceMap[row[0]] = row[7]

		}
	}
	return nil
}

func (p *Pipeline) Process() error {
	var err error
	err = p.extractImageFromMediaSource()
	if err != nil {
		return err
	}
	err = p.extractBasicInfoFromSource()
	if err != nil {
		return err
	}

	err = p.extractPriceInfoFromSource()
	if err != nil {
		return err
	}

	err = p.extractExcludeIDFromSource()

	if err != nil {
		return err
	}

	var datasource = make(map[string]*OutputItemRow)
	for _, s := range p.SalesInfoSource {
		rows, err := s.Rows("Sheet1")
		if err != nil {
			return err
		}
		var curr int
		for rows.Next() {
			row, err := rows.Columns()
			if err != nil {
				return err
			}
			if curr < 4 {
				curr++
				continue
			}

			if len(row) == 0 {
				continue
			}

			if datasource[row[0]] == nil {
				if _, ok := p.excludeProductID[row[0]]; ok {
					continue
				}
				datasource[row[0]] = &OutputItemRow{
					ID:          row[0],
					Name:        fmt.Sprintf("%v %v", row[1], row[3]),
					ImageURLs:   p.imageURLMap[row[0]],
					Description: p.descriptionMap[row[0]],
					Price:       p.priceMap[row[0]],
					Quantity:    row[7],
					Colors:      p.productColorOptionImageURLMap[row[0]],
				}
			}
			curr++
		}
	}

	var row int
	var group int
	for _, data := range datasource {
		for _, v := range data.Colors {
			p.SetCellValueToOutput(fmt.Sprintf("C%v", row), data.Name)
			for imageIdx, imageURL := range data.ImageURLs {
				p.SetCellValueToOutput(fmt.Sprintf("%v%v", imageColumns[imageIdx], row), imageURL)
			}
			p.SetCellValueToOutput(fmt.Sprintf("Q%v", row), "No brand/DD good")
			p.SetCellValueToOutput(fmt.Sprintf("X%v", row), data.Description)
			p.SetCellValueToOutput(fmt.Sprintf("AD%v", row), "0.5")
			p.SetCellValueToOutput(fmt.Sprintf("AE%v", row), "20")
			p.SetCellValueToOutput(fmt.Sprintf("AF%v", row), "25")
			p.SetCellValueToOutput(fmt.Sprintf("AG%v", row), "2")
			p.SetCellValueToOutput(fmt.Sprintf("AW%v", row), data.Quantity)
			p.SetCellValueToOutput(fmt.Sprintf("BA%v", row), data.Price)
			p.SetCellValueToOutput(fmt.Sprintf("A%v", row), group)
			p.SetCellValueToOutput(fmt.Sprintf("AI%v", row), v.ID)
			p.SetCellValueToOutput(fmt.Sprintf("AN%v", row), v.ImageURL)
			row++
		}
		group++
	}

	return nil
}

func (p *Pipeline) Write(writer io.Writer) error {
	return p.output.Write(writer)
}
