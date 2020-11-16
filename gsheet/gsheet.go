package gsheet

import (
	"context"
	"fmt"
	"net/http"
	"strings"

	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

type GSheetCreator struct {
	sheetService *sheets.Service
	driveService *drive.Service

	Colors              map[string]*sheets.Color
	WrapStrategy        string
	VerticalAlignment   string
	HorizontalAlignment string
}

// NewGSheetCreator Take a setup and authenticated client and return a GSheetCreator instance
func NewGSheetCreator(client *http.Client) (*GSheetCreator, error) {
	ctx := context.Background()
	sheetService, err := sheets.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return nil, fmt.Errorf("Unable to retrieve Sheets client: %w", err)
	}

	driveService, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return nil, fmt.Errorf("Unable to retrieve Sheets client: %w", err)
	}
	colorMap := map[string]*sheets.Color{
		"Green":        {Red: 217 / 255.0, Green: 234 / 255.0, Blue: 211 / 255.0},
		"Pink":         {Red: 234 / 255.0, Green: 209 / 255.0, Blue: 220 / 255.0},
		"Purple":       {Red: 217 / 255.0, Green: 210 / 255.0, Blue: 233 / 255.0},
		"Blue":         {Red: 207 / 255.0, Green: 226 / 255.0, Blue: 243 / 255.0},
		"Green2":       {Red: 217 / 255.0, Green: 234 / 255.0, Blue: 211 / 255.0},
		"Yellow":       {Red: 255 / 255.0, Green: 255 / 255.0, Blue: 0 / 255.0},
		"Light Orange": {Red: 255 / 255.0, Green: 202 / 255.0, Blue: 151 / 255.0},
		"Light Cyan":   {Red: 205 / 255.0, Green: 225 / 255.0, Blue: 228 / 255.0},
		"Light Red":    {Red: 255 / 255.0, Green: 200 / 255.0, Blue: 202 / 255.0},
		"Light Yellow": {Red: 255 / 255.0, Green: 244 / 255.0, Blue: 202 / 255.0},
	}

	return &GSheetCreator{
		Colors:              colorMap,
		WrapStrategy:        "WRAP",
		VerticalAlignment:   "MIDDLE",
		HorizontalAlignment: "LEFT",
		sheetService:        sheetService,
		driveService:        driveService}, nil
}

// NewSheetFromStringGrid takes [][]string and returns a formatted Sheet
func (gs *GSheetCreator) NewSheetFromStringGrid(title string, data [][]string, sheetID int64) *sheets.Sheet {
	gridData := &sheets.GridData{}
	rowCount := len(data)
	columnCount := 0
	gridData.RowData = make([]*sheets.RowData, rowCount)
	for i, d := range data {
		if len(d) > columnCount {
			columnCount = len(d)
		}
		gridData.RowData[i] = &sheets.RowData{Values: gs.NewStringRow(d)}
	}

	sh := &sheets.Sheet{
		Properties: &sheets.SheetProperties{
			Title: title,
			GridProperties: &sheets.GridProperties{
				RowCount:    int64(rowCount),
				ColumnCount: int64(columnCount),
			},
		},
		Data: []*sheets.GridData{gridData},
	}

	if sheetID > 0 {
		sh.Properties.SheetId = sheetID
	}

	return sh
}

// NewSheetFromCellGrid takes [][]*sheets.CellData and returns a formatted Sheet
func (gs *GSheetCreator) NewSheetFromCellGrid(title string, data [][]*sheets.CellData, sheetID int64) *sheets.Sheet {
	gridData := &sheets.GridData{}
	rowCount := len(data)
	columnCount := 0
	gridData.RowData = make([]*sheets.RowData, rowCount)
	for i, d := range data {
		if len(d) > columnCount {
			columnCount = len(d)
		}
		gridData.RowData[i] = &sheets.RowData{Values: make([]*sheets.CellData, len(d))}
		for j, cell := range d {
			gridData.RowData[i].Values[j] = cell
		}
	}

	sh := &sheets.Sheet{
		Properties: &sheets.SheetProperties{
			Title: title,
			GridProperties: &sheets.GridProperties{
				RowCount:    int64(rowCount),
				ColumnCount: int64(columnCount),
			},
		},
		Data: []*sheets.GridData{gridData},
	}

	if sheetID > 0 {
		sh.Properties.SheetId = sheetID
	}

	return sh
}

func (gs *GSheetCreator) CreateSpreadsheet(title string, allSheets []*sheets.Sheet) (*sheets.Spreadsheet, error) {
	spread := &sheets.Spreadsheet{
		Properties: &sheets.SpreadsheetProperties{
			Title: title,
		},
		Sheets: allSheets,
	}
	return gs.sheetService.Spreadsheets.Create(spread).Do()
}

func (gs *GSheetCreator) NewStringCell(value string) *sheets.CellData {
	return gs.NewStringCellWithBackground(value, nil)
}

func (gs *GSheetCreator) NewStringCellWithBackground(value string, color *sheets.Color) *sheets.CellData {
	return &sheets.CellData{
		UserEnteredValue: &sheets.ExtendedValue{StringValue: &value},
		UserEnteredFormat: &sheets.CellFormat{
			WrapStrategy:        gs.WrapStrategy,
			VerticalAlignment:   gs.VerticalAlignment,
			HorizontalAlignment: gs.HorizontalAlignment,
			BackgroundColor:     color,
		},
	}
}

func (gs *GSheetCreator) NewNumberCell(value float64) *sheets.CellData {
	return gs.NewNumberCellWithBackground(value, nil)
}

func (gs *GSheetCreator) NewNumberCellWithBackground(value float64, color *sheets.Color) *sheets.CellData {
	return &sheets.CellData{
		UserEnteredValue: &sheets.ExtendedValue{NumberValue: &value},
		UserEnteredFormat: &sheets.CellFormat{
			WrapStrategy:        gs.WrapStrategy,
			VerticalAlignment:   gs.VerticalAlignment,
			HorizontalAlignment: gs.HorizontalAlignment,
			BackgroundColor:     color,
		},
	}
}

func (gs *GSheetCreator) NewBoolCell(value bool) *sheets.CellData {
	return gs.NewBoolCellWithBackground(value, nil)
}

func (gs *GSheetCreator) NewBoolCellWithBackground(value bool, color *sheets.Color) *sheets.CellData {
	return &sheets.CellData{
		UserEnteredValue: &sheets.ExtendedValue{BoolValue: &value},
		UserEnteredFormat: &sheets.CellFormat{
			WrapStrategy:        gs.WrapStrategy,
			VerticalAlignment:   gs.VerticalAlignment,
			HorizontalAlignment: gs.HorizontalAlignment,
			BackgroundColor:     color,
		},
	}
}

func (gs *GSheetCreator) NewStringRow(data []string) []*sheets.CellData {
	row := make([]*sheets.CellData, len(data))
	for i := range data {
		row[i] = gs.NewStringCellWithBackground(data[i], nil)
	}
	return row
}

func (gs *GSheetCreator) NewStringRowWithBackground(data []string, color *sheets.Color) []*sheets.CellData {
	row := make([]*sheets.CellData, len(data))
	for i := range data {
		row[i] = gs.NewStringCellWithBackground(data[i], color)
	}
	return row
}

func (gs *GSheetCreator) HideColumns(sheetID, columnIndexStart, columnIndexEnd int64) *sheets.Request {
	return &sheets.Request{
		UpdateDimensionProperties: &sheets.UpdateDimensionPropertiesRequest{
			Range: &sheets.DimensionRange{
				SheetId:    sheetID,
				Dimension:  "COLUMNS",
				StartIndex: columnIndexStart,
				EndIndex:   columnIndexEnd,
			},
			Properties: &sheets.DimensionProperties{HiddenByUser: true},
			Fields:     "hiddenByUser",
		},
	}
}

func (gs *GSheetCreator) FreezeRows(sheetID, frozenNumber int64) *sheets.Request {
	return &sheets.Request{
		UpdateSheetProperties: &sheets.UpdateSheetPropertiesRequest{
			Properties: &sheets.SheetProperties{SheetId: sheetID,
				GridProperties: &sheets.GridProperties{FrozenRowCount: frozenNumber},
			},
			Fields: "gridProperties.frozenRowCount",
		},
	}
}

func (gs *GSheetCreator) UpdateColumnWidth(sheetID, columnIndexStart, columnIndexEnd, width int64) *sheets.Request {
	return &sheets.Request{
		UpdateDimensionProperties: &sheets.UpdateDimensionPropertiesRequest{
			Range: &sheets.DimensionRange{
				SheetId:    sheetID,
				Dimension:  "COLUMNS",
				StartIndex: columnIndexStart,
				EndIndex:   columnIndexEnd,
			},
			Properties: &sheets.DimensionProperties{
				PixelSize: width,
			},
			Fields: "pixelSize",
		},
	}
}

func (gs *GSheetCreator) UpdateRowHeight(sheetID, rowIndexStart, rowIndexEnd, height int64) *sheets.Request {
	return &sheets.Request{
		UpdateDimensionProperties: &sheets.UpdateDimensionPropertiesRequest{
			Range: &sheets.DimensionRange{
				SheetId:    sheetID,
				Dimension:  "ROWS",
				StartIndex: rowIndexStart,
				EndIndex:   rowIndexEnd,
			},
			Properties: &sheets.DimensionProperties{
				PixelSize: height,
			},
			Fields: "pixelSize",
		},
	}
}

func (gs *GSheetCreator) UpdateCellBackground(gridRange *sheets.GridRange, color *sheets.Color) *sheets.Request {
	return &sheets.Request{
		RepeatCell: &sheets.RepeatCellRequest{
			Range: gridRange,
			Cell: &sheets.CellData{
				UserEnteredFormat: &sheets.CellFormat{
					BackgroundColor: color,
				},
			},
			Fields: "userEnteredFormat(backgroundColor)",
		},
	}
}

func (gs *GSheetCreator) ExecuteUpdates(spreadsheetId string, updates []*sheets.Request) (*sheets.BatchUpdateSpreadsheetResponse, error) {
	req := &sheets.BatchUpdateSpreadsheetRequest{Requests: updates}
	return gs.sheetService.Spreadsheets.BatchUpdate(spreadsheetId, req).Do()
}

func (gs *GSheetCreator) ShareWithDomain(fileID, domain string) (*drive.Permission, error) {
	return gs.driveService.Permissions.Create(fileID, &drive.Permission{
		Type:   "domain",
		Role:   "writer",
		Domain: domain,
	}).Do()
}

func (gs *GSheetCreator) MoveToFolder(fileID, folderID string) (*drive.File, error) {
	f, err := gs.driveService.Files.Get(fileID).Do()
	if err != nil {
		return nil, err
	}

	return gs.driveService.Files.Update(fileID, &drive.File{}).
		AddParents(folderID).
		RemoveParents(strings.Join(f.Parents, ",")).
		Fields("id, parents").
		Do()
}
