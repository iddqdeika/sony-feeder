package main

import (
	"encoding/xml"
	"fmt"
	excel "github.com/360EntSecGroup-Skylar/excelize"
	"io/ioutil"
	"os"
	"reflect"
)

const (
	fileName = "sony.xml"
)

func main() {
	data, err := ioutil.ReadFile(fileName)
	if err != nil {
		fmt.Println("чтобы распарсить файл - назовите sony.xml и положите в текущую папку.")
		exit("не получается считать файл: " + err.Error())
	}
	f := new(Feed)
	err = xml.Unmarshal(data, &f)
	if err != nil {
		exit("не получается распарсить xml: " + err.Error())
	}

	res := excel.NewFile()

	fm := make(map[string]int)

	res.NewSheet("items")
	for rn, e := range f.Entry {
		v := reflect.ValueOf(e)
		for fn := 0; fn < v.NumField(); fn++ {
			fv := v.Field(fn)
			name := v.Type().Field(fn).Name
			value := fv.String()
			if _, ok := fm[name]; !ok {
				fm[name] = len(fm) + 1
				axis, err := excel.CoordinatesToCellName(fm[name], 1)
				if err != nil {
					exit("не получается сгенерить excel файл: " + err.Error())
				}
				err = res.SetCellValue("items", axis, name)
				if err != nil {
					exit("не получается сгенерить excel файл: " + err.Error())
				}
			}
			cn := fm[name]
			axis, err := excel.CoordinatesToCellName(cn, rn+2)
			if err != nil {
				exit("не получается сгенерить excel файл: " + err.Error())
			}
			err = res.SetCellValue("items", axis, value)
			if err != nil {
				exit("не получается сгенерить excel файл: " + err.Error())
			}
		}
	}
	err = res.SaveAs("sony_feed_" + f.Updated[:10] + ".xlsx")
	if err != nil {
		panic("не получается сгенерить excel файл: " + err.Error())
	}
	exit("всё ок")
}

func exit(s string) {
	fmt.Println(s)
	fmt.Println("нажмите ENTER чтобы выйти")
	fmt.Scanln()
	os.Exit(0)
}

type Feed struct {
	Title   string  `xml:"title"`
	Updated string  `xml:"updated"`
	Entry   []Entry `xml:"entry"`
}

type Entry struct {
	Id           string `xml:"id"`
	Title        string `xml:"title"`
	Description  string `xml:"description"`
	Link         string `xml:"link"`
	Image        string `xml:"image_link"`
	Condition    string `xml:"condition"`
	Availability string `xml:"availability"`
	Price        string `xml:"price"`
	Gtin         string `xml:"gtin"`
	Mpn          string `xml:"mpn"`
	ProductType  string `xml:"product_type"`
}
