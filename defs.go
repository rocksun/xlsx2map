package xlsx2map

type XlsxFileDef struct {
	SheetDefs []*SheetDef `json:"sheets"`
}

func (file *XlsxFileDef) GetSheetDef(name string) *SheetDef {
	for _, sheetDef := range file.SheetDefs {
		if sheetDef.ValidAlias(name) {
			return sheetDef
		}
	}
	return nil
}

type SheetDef struct {
	Key       string      `json:"key"`
	Aliases   []string    `json:"aliases"`
	FieldDefs []*FieldDef `json:"fields"`
}

func (sheetDef *SheetDef) GetTitle() string {
	if len(sheetDef.Aliases) >= 1 {
		return sheetDef.Aliases[0]
	}
	return sheetDef.Key
}

func (sheetDef *SheetDef) ValidAlias(name string) bool {
	for _, alias := range sheetDef.Aliases {
		if alias == name {
			return true
		}
	}
	return false
}

func (sheetDef *SheetDef) GetFieldDef(name string) *FieldDef {
	for _, fieldDef := range sheetDef.FieldDefs {
		if fieldDef.ValidAlias(name) {
			return fieldDef
		}
	}
	return nil
}

type FieldDef struct {
	Key        string   `json:"key"`
	Index      int      `json:"index"`
	Aliases    []string `json:"aliases"`
	DataType   string   `json:"dataType"`
	DataOption string   `json:"dataOption"`
}

func (fieldDef *FieldDef) ValidAlias(name string) bool {
	for _, alias := range fieldDef.Aliases {
		if alias == name {
			return true
		}
	}
	return false
}

func (fieldDef *FieldDef) ParseValue(valueStr string) (interface{}, error) {
	parseFunc := ParseDataFuncs.Get(fieldDef.DataType)
	return parseFunc(valueStr, fieldDef)
}

func (fieldDef *FieldDef) GetTitle() string {
	if len(fieldDef.Aliases) >= 1 {
		return fieldDef.Aliases[0]
	}
	return fieldDef.Key
}

type Columns struct {
	FieldDefs map[int]*FieldDef
}

func (columns *Columns) GetFieldDef(index int) *FieldDef {
	if fdef := columns.FieldDefs[index]; fdef != nil {
		return fdef
	}
	return nil
}

func (columns *Columns) AddColumns(index int, fieldDef *FieldDef) {
	columns.FieldDefs[index] = fieldDef
}
