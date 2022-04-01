package xlsx2map

import (
	"strconv"

	"github.com/xuri/excelize/v2"
)

type ParseDataFuncFactory struct {
	ParseDataFuncs map[string]ParseDataFunc
}

func GettParseDataFuncFactory() *ParseDataFuncFactory {
	factory := &ParseDataFuncFactory{make(map[string]ParseDataFunc)}
	factory.AddFunc("string", ParseString)
	factory.AddFunc("int", ParseInt)
	factory.AddFunc("float", ParseFloat)
	factory.AddFunc("ExcelDate", ParseExcelDate)
	return factory
}

func (factory *ParseDataFuncFactory) Get(dataType string) ParseDataFunc {
	if dataType == "" {
		return factory.ParseDataFuncs["string"]
	}
	return factory.ParseDataFuncs[dataType]
}

func (factory *ParseDataFuncFactory) AddFunc(dataType string, pdFunc ParseDataFunc) {
	factory.ParseDataFuncs[dataType] = pdFunc
}

var ParseDataFuncs *ParseDataFuncFactory

func init() {
	ParseDataFuncs = GettParseDataFuncFactory()
}

type ParseDataFunc func(valueStr string, ops interface{}) (interface{}, error)

func ParseString(valueStr string, ops interface{}) (interface{}, error) {
	return valueStr, nil
}

func ParseInt(valueStr string, ops interface{}) (interface{}, error) {
	intValue, err := strconv.ParseInt(valueStr, 0, 64)
	if err != nil {
		return nil, err
	}
	return intValue, nil
}

func ParseFloat(valueStr string, ops interface{}) (interface{}, error) {
	floatValue, err := strconv.ParseFloat(valueStr, 64)
	if err != nil {
		return nil, err
	}
	return floatValue, nil
}

func ParseExcelDate(valueStr string, ops interface{}) (interface{}, error) {
	excelDate, err := strconv.ParseFloat(valueStr, 64)
	if err != nil {
		return nil, err
	}
	excelTime, err := excelize.ExcelDateToTime(excelDate, false)
	if err != nil {
		return nil, err
	}
	return excelTime, nil
}
