package main

import (
	"github.com/tealeg/xlsx"
	"os"
	"strings"
	"path/filepath"
	"log"
	"time"
	"encoding/json"
	"io/ioutil"
	"unicode"
	"strconv"
	"runtime"
	"flag"
)
var f string

func init()  {
	flag.StringVar(&f,"f", "game.xlsx", "Input excel config file")
}
func main() {
	flag.Parse()
	if !strings.Contains(f,".xlsx") {
		log.Printf("only support *.xlsx files.")
		waitAndExit(-1,3 )
	}
	log.Printf("flag->%s\n",f)
	currentDir, err := filepath.Abs(filepath.Dir(os.Args[0]))
	config := make(map[string][]map[string]interface{})
	if !strings.Contains(f,"/") && !strings.Contains(f, "\\") {
		f = currentDir + "/" + f
	}
	f = strings.Replace(f, "\\", "/", -1)
	log.Printf("File in %s", f)
	xlFile, err := xlsx.OpenFile(f)
	if err != nil {
		log.Printf("open failed: %s\n", err)
		waitAndExit(-1,3 )
	}
	for _, sheet := range xlFile.Sheets {
		// 配置数据
		sheetList := make([]map[string]interface{},0)
		// keys
		var keys []string
		for j, row := range sheet.Rows {
			// Sheet中文标题
			if j == 0 {
				continue
			}
			if j == 1{
				keys = make([]string,0)
				for _, value := range row.Cells {
					// 去除空格
					value := strings.TrimSpace(value.String())
					// 空值处理
					//value = strings.Replace(value, "", "null", -1)
					if value != "" {
						keys = append(keys, value)
					}
				}
				continue
			}
			// 数据
			rowMap := make(map[string]interface{})
			for ij, value := range row.Cells {
				// 去除空格
				value := strings.TrimSpace(value.String())
				// 空值处理
				//value = strings.Replace(value, "", "null", -1)
				maxIndex := len(keys) -1
				if ij > maxIndex {
					break
				}
				rowMap[keys[ij]] = parseCellValue(value)
			}
			if len(rowMap) < 1 {
				continue
			}
			// 加入每行数据
			sheetList = append(sheetList, rowMap)
		}
		config[sheet.Name] = sheetList
	}
	log.Println("----------------------------------OK, file handle complete----------------------------------------")
	configJson, err := json.Marshal(config)
	if err != nil {
		log.Println("Parse config file error ，", err)
	}
	log.Println(string(configJson))
	// 写文件JSON
	if err != nil {
		log.Println("Get current dir error, ", err)
	}
	// Json文件名
	jsonFileName := strings.Replace(f,".xlsx",".json",1)
	err = ioutil.WriteFile(jsonFileName, configJson, 0666)
	if err != nil {
		log.Println("Write JSON file error, ", err)
	}else {
		log.Println("JSON file at " + jsonFileName)
	}
	log.Println("Parse excel file success, done. ")
	log.Println("策划棒棒哒，解析成功！")
	waitAndExit(0,3 )
}

func waitAndExit(code, dur int) {
	log.Println("Exit after " + strconv.Itoa(dur) + " seconds ...")
	time.Sleep(time.Second * time.Duration(dur))
	os.Exit(code)
}

// 判断是否是数字
func stringIsDigit(value string) bool  {
	for _, ch := range  value{
		if !unicode.IsDigit(ch) {
			return  false
		}
	}
	return true
}
func checkErr(err error)  {
	if err != nil {
		log.Fatalf("Parse error , %v", err)
	}
}

func parseCellValue(value string)interface{}  {
	var cellValue interface{}
	var err error
	// 处理不同类型的数值
	if strings.Contains(value,"."){// 浮点数
		cellValue, err = strconv.ParseFloat(value, getBit())
		checkErr(err)
	}else if strings.Contains(value,"|"){ // 数组
		cellValue =  parseArray(value)
	}else if stringIsDigit(value) && len(value) > 0 {// 整数
		cellValue, err = strconv.ParseInt(value, 10, getBit())
		checkErr(err)
	}else {// 字符串
		cellValue = value
	}
	return cellValue
}

func parseArray(s string) interface{} {
	var array = make([]interface{},0)
	sArray := strings.Split(s,"|")
	for _, v := range sArray {
		array =  append(array, parseCellValue(v))
	}
	return array
}

func getBit() int {
	arch := runtime.GOARCH
	var bit int
	if strings.Contains(arch, "64") {
		bit = 64
	}else {
		bit = 32
	}
	return bit
}
