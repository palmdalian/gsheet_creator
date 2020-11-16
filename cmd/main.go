package main

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"net/http"
	"os"

	"golang.org/x/net/context"
	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/sheets/v4"

	"github.com/palmdalian/gsheet_creator/gsheet"
)

// Retrieve a token, saves the token, then returns the generated client.
func getClient(config *oauth2.Config) *http.Client {
	// The file token.json stores the user's access and refresh tokens, and is
	// created automatically when the authorization flow completes for the first
	// time.
	tokFile := "token.json"
	tok, err := tokenFromFile(tokFile)
	if err != nil {
		tok = getTokenFromWeb(config)
		saveToken(tokFile, tok)
	}
	return config.Client(context.Background(), tok)
}

// Request a token from the web, then returns the retrieved token.
func getTokenFromWeb(config *oauth2.Config) *oauth2.Token {
	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
	fmt.Printf("Go to the following link in your browser then type the "+
		"authorization code: \n%v\n", authURL)

	var authCode string
	if _, err := fmt.Scan(&authCode); err != nil {
		log.Fatalf("Unable to read authorization code: %v", err)
	}

	tok, err := config.Exchange(context.TODO(), authCode)
	if err != nil {
		log.Fatalf("Unable to retrieve token from web: %v", err)
	}
	return tok
}

// Retrieves a token from a local file.
func tokenFromFile(file string) (*oauth2.Token, error) {
	f, err := os.Open(file)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	tok := &oauth2.Token{}
	err = json.NewDecoder(f).Decode(tok)
	return tok, err
}

// Saves a token to a file path.
func saveToken(path string, token *oauth2.Token) {
	fmt.Printf("Saving credential file to: %s\n", path)
	f, err := os.OpenFile(path, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
	if err != nil {
		log.Fatalf("Unable to cache oauth token: %v", err)
	}
	defer f.Close()
	json.NewEncoder(f).Encode(token)
}

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
	client := getClient(config)

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
