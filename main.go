// tradelist project main.go
package main

import (
	//	"time"
	//	"fmt"
	"log"

	"strconv"

	"os"

	"io"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/lxn/walk"
	. "github.com/lxn/walk/declarative"
)

type mymainwindow struct {
	*walk.MainWindow
	Name        *walk.LineEdit
	Lineweight  *walk.LineEdit
	Price       *walk.LineEdit
	Other       *walk.LineEdit
	Allweight   *walk.LineEdit
	Totalcost   *walk.LineEdit
	Totalweight *walk.LineEdit
	Tp          [30]*walk.ComboBox
	Num         [30]*walk.LineEdit
	len         [30]*walk.LineEdit
	weight      [30]*walk.LineEdit
	Generate    *walk.PushButton
	calculate   *walk.PushButton
}

var m = map[string]float64{
	"Φ10": 0.617,
	"Φ12": 0.89,
	"Φ14": 1.21,
	"Φ16": 1.58,
	"Φ18": 2.0,
	"Φ20": 2.47,
	"Φ22": 2.98,
	"Φ25": 3.85,
}
var mn *walk.Menu

func main() {
	mw := new(mymainwindow)
	mymw := MainWindow{
		AssignTo: &mw.MainWindow,
		Title:    "清单生成器",
		MinSize:  Size{500, 500},
		Layout:   VBox{},
		Children: []Widget{
			Composite{
				MaxSize: Size{0, 40},
				Layout:  HBox{},
				Children: []Widget{
					Label{Text: "客户"},
					LineEdit{AssignTo: &mw.Name},
					PushButton{AssignTo: &mw.calculate, Text: "计算"},
					PushButton{AssignTo: &mw.Generate, Text: "生成清单"},
				},
			},
			GroupBox{

				Layout: Grid{Columns: 4},
				Children: []Widget{
					Label{Text: "线材重量（公斤）"},
					LineEdit{AssignTo: &mw.Lineweight},
					Label{Text: "螺纹总重（公斤）"},
					LineEdit{AssignTo: &mw.Allweight, ReadOnly: true},

					Label{Text: "总重量（公斤）"},
					LineEdit{AssignTo: &mw.Totalweight, ReadOnly: true},
					Label{Text: "价格（元/吨）"},
					LineEdit{AssignTo: &mw.Price},

					Label{Text: "其他耗材（元）"},
					LineEdit{AssignTo: &mw.Other},
					Label{Text: "总金额（元）"},
					LineEdit{AssignTo: &mw.Totalcost, ReadOnly: true},
				},
			},
		},
	}

	c := Composite{
		MinSize:  Size{0, 50},
		Layout:   HBox{},
		Children: []Widget{},
	}

	box := GroupBox{
		Layout:   Grid{Columns: 5},
		Children: []Widget{},
	}
	box.Children = append(box.Children, Label{Text: "序号"}, Label{Text: "型号"}, Label{Text: "长度（米）"}, Label{Text: "数量（支）"}, Label{Text: "重量（公斤）"})
	for i := 0; i < 30; i++ {

		l := Label{Text: strconv.Itoa(i + 1)}
		tp := ComboBox{
			//Editable: true,
			//Value:    Bind("PreferredFood"),

			AssignTo: &mw.Tp[i],
			MinSize:  Size{60, 0},
			Model:    []string{"", "Φ10", "Φ12", "Φ14", "Φ16", "Φ18", "Φ20", "Φ22", "Φ25"},
		}
		len := LineEdit{
			AssignTo: &mw.len[i],
		}
		n := LineEdit{
			AssignTo: &mw.Num[i],
		}
		weight := LineEdit{
			AssignTo: &mw.weight[i],
			ReadOnly: true,
		}

		box.Children = append(box.Children, l, tp, len, n, weight)
	}
	c.Children = append(c.Children, box)
	mymw.Children = append(mymw.Children, c)
	if err := mymw.Create(); err != nil {
		log.Fatalln(err)
	}
	mw.calculate.Clicked().Attach(func() {
		go func() {
			allweight := 0.0
			for i := 0; i < 30; i++ {

				weight := m[mw.Tp[i].Text()] * Atof(mw.Num[i].Text()) * Atof(mw.len[i].Text())
				allweight += weight
				mw.weight[i].SetText(Ftoa(weight))
			}
			mw.Allweight.SetText(Ftoa(allweight))
			lineweight := Atof(mw.Lineweight.Text())
			price := Atof(mw.Price.Text())
			other := Atof(mw.Other.Text())
			totalweight := lineweight + allweight
			mw.Totalweight.SetText(Ftoa(totalweight))
			totalcost := totalweight*price/1000 + other
			mw.Totalcost.SetText(Ftoa(totalcost))

		}()
	})
	mw.Generate.Clicked().Attach(func() {
		go func() {
			/////
			temp, e := os.Open("模版/template.xlsx")
			if e != nil {
				panic(e)
			}
			defer temp.Close()
			newfile, e := os.Create(mw.Name.Text() + "_" + mw.Totalcost.Text() + ".xlsx")
			if e != nil {
				panic(e)
			}
			defer newfile.Close()
			io.Copy(newfile, temp)
			file, e := excelize.OpenFile(mw.Name.Text() + "_" + mw.Totalcost.Text() + ".xlsx")
			if e != nil {
				panic(e)
			}
			file.SetCellValue("sheet1", "B3", mw.Name.Text())
			file.SetCellValue("sheet1", "C39", Atof(mw.Lineweight.Text()))
			file.SetCellValue("sheet1", "C41", Atof(mw.Price.Text()))
			file.SetCellValue("sheet1", "C42", Atof(mw.Other.Text()))
			for i := 0; i < 30; i++ {
				if mw.Tp[i].Text() != "" {
					//file.SetCellValue("sheet1", "A"+strconv.Itoa(6+i), strconv.Itoa(i+1))
					file.SetCellValue("sheet1", "B"+strconv.Itoa(6+i), mw.Tp[i].Text())
					file.SetCellValue("sheet1", "C"+strconv.Itoa(6+i), Atof(mw.len[i].Text()))
					file.SetCellValue("sheet1", "D"+strconv.Itoa(6+i), Atof(mw.Num[i].Text()))
					file.SetCellValue("sheet1", "E"+strconv.Itoa(6+i), Atof(mw.weight[i].Text()))
				}
			}

			file.UpdateLinkedValue()
			file.Save()
			////
		}()
	})
	mw.Run()

}
func Atof(a string) float64 {
	f, er := strconv.ParseFloat(a, 64)
	if er != nil {

		return 0.0
	}
	return f
}
func Ftoa(f float64) string {
	a := strconv.FormatFloat(f, 'f', 2, 64)
	return a
}
