package main

import (
	"log"
	"net/smtp"
	"strings"
	"xlsx"
)

var formattedrowdata []string

func main() {
	Printexcelfile()
}

//Printexcelfile send the contents of an Excel file
//In the excel file, e-mails should be in column A
//In the excel file, the subject line should be column B
//The data from your spreadsheet(s) will be send as individual emails, starting with Row 2
//Great for accountants, teachers, sales managers, or anyone that manages multiple streams of users & their data with excel

func Printexcelfile() {
	excelFileName := "xxxxxxxxxxx" //path to excel file
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		println(err)
	}
	for _, sheet := range xlFile.Sheets {
		for x, row := range sheet.Rows {
			if x > 0 {
				for i := range row.Cells {
					column := sheet.Cell(0, i)
					columnvalue := column.String()
					cell := sheet.Cell(x, i)
					cellvalue := cell.String()
					newformattedcell := columnvalue + " : " + cellvalue + "\n"
					formattedrowdata = append(formattedrowdata, newformattedcell)

				}

				recipient := (sheet.Cell(x, 0)).String()
				subjectline := (sheet.Cell(x, 1)).String()
				Send(recipient, subjectline, formattedrowdata)
				formattedrowdata = nil
			}
		}
	}
}

//Send function sends an email using smtp
func Send(receiver string, subject string, bodydata []string) {
	from := "xxxxxxx" //sender's gmail address
	pass := "xxxxxxx" //this is not your regular password.  A special "app password" Must be requested thru your google account to bypass 2 step verificaton
	to := receiver
	body := strings.Join(bodydata[2:], "")
	print(body)
	msg := []byte("From: " + from + "\n" +
		"To: " + to + "\n" +
		"Subject:" + subject + "\n\n" + body)

	err := smtp.SendMail("smtp.gmail.com:587",
		smtp.PlainAuth("", from, pass, "smtp.gmail.com"),
		from, []string{to}, msg)

	if err != nil {
		log.Printf("smtp error: %s", err)
		return
	}
}
