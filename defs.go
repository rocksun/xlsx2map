package xlsx2map

type XlsxFileDef struct {
	SheetDefs []SheetDef `json:"sheets"`
}

type SheetDef struct {
	Key       string     `json:"key"`
	Aliases   []string   `json:"aliases"`
	FieldDefs []FieldDef `json:"fields"`
}

type FieldDef struct {
	Key     string   `json:"key"`
	Index   int      `json:"index"`
	Aliases []string `json:"aliases"`
}
