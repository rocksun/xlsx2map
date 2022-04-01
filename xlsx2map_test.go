package xlsx2map

import (
	"fmt"
	"os"
	"testing"
)

func TestUnmarshal(t *testing.T) {

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

	xlsxFile := "sample_file.xlsx"

	got, err := Unmarshal(xlsxFile, def, nil)
	if err != nil {
		t.Errorf("Unmarshal() error = %v, wantErr %v", err, nil)
		return
	}
	fmt.Println(got)
	// if !reflect.DeepEqual(got, tt.want) {
	// 	t.Errorf("Unmarshal() = %v, want %v", got, tt.want)
	// }

}
