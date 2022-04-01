package xlsx2map

import (
	"fmt"
	"os"
	"testing"
)

func TestXlsxFileDef(t *testing.T) {
	def := &XlsxFileDef{}
	file, err := os.Open("sample_def.json")

	if err != nil {
		t.Errorf("expected no err, but got %v", err)
	}

	defer file.Close()

	loadErr := LoadXlsxFileDef(file, def)
	if loadErr != nil {
		t.Errorf("expected no err, but got %v", loadErr)
	}
	expected := "visitors"
	fmt.Println(def)
	if def.SheetDefs[0].Key != expected {
		t.Errorf("expected %v, but got %v", expected, def.SheetDefs[0].Key)
	}

	key1 := "name"
	if def.SheetDefs[0].FieldDefs[0].Key != key1 {
		t.Errorf("expected %v, but got %v", key1, def.SheetDefs[0].FieldDefs[0].Key)
	}

	nilSheetDef := def.GetSheetDef("xxx")
	if nilSheetDef != nil {
		t.Errorf("expected return nil, but got %v", nilSheetDef)
	}

	realSheetDef := def.GetSheetDef("Visitors List")
	if realSheetDef.Key != "visitors" {
		t.Errorf("expected return visitors, but got %v", realSheetDef.Key)
	}
}
