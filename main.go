package main

import (
	"encoding/json"
	"fmt"
	"net/http"
	"strconv"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

type DataResponse struct {
	Power        string
	Efficiency   string
	FileLocation string
}

type ErrorMessage struct {
	Error   string
	Message string
}

func main() {
	http.HandleFunc("/", handleRequest)
	http.ListenAndServe(":3000", nil)
}

func handleRequest(w http.ResponseWriter, r *http.Request) {

	pressure := r.URL.Query().Get("pressure")
	flow := r.URL.Query().Get("flow")
	speed := r.URL.Query().Get("speed")
	temperature := r.URL.Query().Get("temp")
	altitude := r.URL.Query().Get("alt")
	fileLocation := r.URL.Query().Get("fl")
	newFileLocation := r.URL.Query().Get("nfl")

	f, err := excelize.OpenFile(fileLocation)
	if err != nil {
		m := "Invalid File Location Provided : " + fileLocation + " error : " + err.Error()
		printErrorResponse(m, w, r)
		return
	}

	if newFileLocation == "" {
		m := "Invalid New File Location Provided : " + newFileLocation
		printErrorResponse(m, w, r)
		return
	}

	//set values
	if pressure != "" {
		prvalue, err := strconv.ParseFloat(pressure, 32)
		if err != nil {
			m := "Invalid Value Pressure" + err.Error()
			printErrorResponse(m, w, r)
			return
		}
		f.SetCellValue("Data", "F5", prvalue)
	}

	if flow != "" {
		flvalue, err := strconv.ParseFloat(flow, 32)
		if err != nil {
			m := "Invalid Value Flow" + err.Error()
			printErrorResponse(m, w, r)
			return
		}
		f.SetCellValue("Data", "F4", flvalue)
	}

	if speed != "" {
		spvalue, err := strconv.ParseFloat(speed, 32)
		if err != nil {
			m := "Invalid Value Speed" + err.Error()
			printErrorResponse(m, w, r)
			return
		}
		f.SetCellValue("Data", "F6", spvalue)
	}

	if temperature != "" {
		temvalue, err := strconv.ParseFloat(temperature, 64)
		if err != nil {
			m := "Invalid Value Temperature" + err.Error()
			printErrorResponse(m, w, r)
			return
		}
		f.SetCellValue("Data", "F7", temvalue)
	}

	if altitude != "" {
		avalue, err := strconv.ParseFloat(altitude, 32)
		if err != nil {
			m := "Invalid Value Pressure" + err.Error()
			printErrorResponse(m, w, r)
			return
		}
		f.SetCellValue("Data", "F8", avalue)
	}

	f.CalcCellValue("Data", "T53")
	power, err := f.GetCellValue("Data", "T53")
	if err != nil {
		m := "Invalid Cell Value Calculated - Power : " + err.Error()
		printErrorResponse(m, w, r)
		return
	}

	f.CalcCellValue("Data", "T54")
	efficiency, err := f.GetCellValue("Data", "T54")
	if err != nil {
		m := "Invalid Cell Value Calculated - Efficiency : " + err.Error()
		printErrorResponse(m, w, r)
		return
	}

	// Get all the rows in the Sheet1.
	tn := time.Now()
	vl := fmt.Sprint(tn.Format("01022006150405"))
	filename := newFileLocation + fmt.Sprint(vl) + ".xlsx"

	//update xls file
	f.UpdateLinkedValue()

	if err := f.SaveAs(filename); err != nil {
		m := "Cannot create new file : " + err.Error()
		printErrorResponse(m, w, r)
		return
	}

	response := DataResponse{power, efficiency, filename}

	js, err := json.Marshal(response)
	if err != nil {
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}

	fmt.Println(response)
	w.Header().Set("Content-Type", "application/json")
	w.Write(js)
}

/*
* Prints Response
 */
func printErrorResponse(message string, w http.ResponseWriter, r *http.Request) {
	m := ErrorMessage{"1", message}
	js, err := json.Marshal(m)
	if err != nil {
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}
	w.Header().Set("Content-Type", "application/json")
	w.Write(js)
}
