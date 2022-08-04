package main

import (
	"fmt"
	"io/ioutil"
	"log"

	"golang.org/x/oauth2/google"
	"google.golang.org/api/sheets/v4"

	"github.com/palmdalian/gsheet_creator/gsheet"
)

func main() {
	b, err := ioutil.ReadFile("../google_client.json")
	if err != nil {
		log.Fatalf("Unable to read client secret file: %v", err)
	}

	// If modifying these scopes, delete your previously saved token.json.
	config, err := google.ConfigFromJSON(b, "https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/presentations")
	if err != nil {
		log.Fatalf("Unable to parse client secret file to config: %v", err)
	}
	client := gsheet.GetClient(config, "token.json")

	gs, err := gsheet.NewGSheetCreator(client)
	if err != nil {
		log.Fatalf("Unable to retrieve Sheets client: %v", err)
	}

	sheetData := [][]string{
		{"heelp", "2", "3", "4", "5", "6", "7", "8"},
		{"What", "is", "the", "deal", "with", "airlines"},
		{"This", "is", "the", "third", "row"},
		{"How", "is", "life", "today"},
		{"ONE", "Hi", "there"},
		{"TWO", "What", "is", "up?"},
	}

	sh := gs.NewSheetFromStringGrid("hello", sheetData, -1)
	resp, err := gs.CreateSpreadsheet("Full Title", []*sheets.Sheet{sh})
	if err != nil {
		fmt.Println(err)
	}
	fmt.Println(resp.SpreadsheetUrl)

	documentID := resp.SpreadsheetId
	sheetID := resp.Sheets[0].Properties.SheetId
	updates := []*sheets.Request{}

	updates = append(updates, gs.UpdateCellBackground(&sheets.GridRange{SheetId: sheetID, StartRowIndex: 0, EndRowIndex: 1}, gs.Colors["Yellow"]))
	updates = append(updates, gs.FreezeRows(sheetID, 1))
	updates = append(updates, gs.UpdateColumnWidth(sheetID, 1, 2, 800))
	updates = append(updates, gs.UpdateRowHeight(sheetID, 1, 2, 550))
	updates = append(updates, gs.HideColumns(sheetID, 3, 4))

	_, err = gs.ExecuteUpdates(documentID, updates)
	if err != nil {
		fmt.Println(err)
	}
	_, err = gs.MoveToFolder(documentID, "1WHpzo93cN1QKCfPkif3mSTdKfWZ998Sj")
	if err != nil {
		fmt.Println(err)
	}
	_, err = gs.ShareWithDomain(documentID, "tonal.com")
	if err != nil {
		fmt.Println(err)
	}

}
