package xlsx2map

import (
	"encoding/json"
	"io"
	"io/ioutil"
)

func LoadXlsxFileDef(in io.Reader, def *XlsxFileDef) error {

	byteValue, ioErr := ioutil.ReadAll(in)
	if ioErr != nil {
		return ioErr
	}

	mErr := json.Unmarshal([]byte(byteValue), def)
	if mErr != nil {
		return mErr
	}

	return nil

}
