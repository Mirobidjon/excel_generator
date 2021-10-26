package excel_generator

import (
	"encoding/json"
	"fmt"
	"os"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/minio/minio-go"
	"github.com/xtgo/uuid"
)

type Data struct {
	Rows []map[string]interface{} `json:"rows"`
}

type Response struct {
	FileName string `json:"file_name"`
	Url      string `json:"url"`
}

type Pair struct {
	First  string
	Second string
}

func NewPair(first, second string) Pair {
	return Pair{
		First:  first,
		Second: second,
	}
}

func generate(data Data, f *excelize.File, titles []Pair) {
	key := 1
	for _, row := range data.Rows {
		key++
		column := 'A'
		f.SetCellValue("Sheet1", fmt.Sprintf("%c%d", 'A', key), key-1)
		for _, v := range titles {
			for k, value := range row {
				if v.First == k {
					column++
					f.SetCellValue("Sheet1", fmt.Sprintf("%c%d", column, key), value)
					break
				}
			}
		}
	}
}

func GenerateExcel(data []byte, bucketName, minioEndpoint, accessKey, secretKey string, titles ...Pair) (*Response, error) {
	var model Data

	err := json.Unmarshal(data, &model.Rows)
	if err != nil {
		return nil, err
	}

	f := excelize.NewFile()

	column := 'A'
	for _, v := range titles {
		f.SetCellValue("Sheet1", fmt.Sprintf("%c%d", column, 1), v.Second)
		column++
	}

	style, _ := f.NewStyle(`{"alignment":{"horizontal":"center", "vertical": "center", "wrap_text": true}}`)
	f.SetCellStyle("Sheet1", "A1", fmt.Sprintf("%c%d", column, len(model.Rows)+1), style)
	f.SetColWidth("Sheet1", "B", string(column), 30)
	f.SetRowHeight("Sheet1", 1, 35)

	generate(model, f, titles)

	v, err := minioUploader(bucketName, minioEndpoint, accessKey, secretKey, f)
	if err != nil {
		return nil, err
	}

	return v, nil
}

func minioUploader(bucketName, minioEndpoint, accessKey, secretKey string, f *excelize.File) (*Response, error) {
	filename := uuid.NewRandom()
	dst, _ := os.Getwd()

	excelContentType := "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

	fmt.Println("generate file name: ", filename)

	minioClient, err := minio.New(minioEndpoint, accessKey, secretKey, false)
	if err != nil {
		return nil, err
	}

	exists, _ := minioClient.BucketExists(bucketName)

	if !exists {
		minioClient.MakeBucket(bucketName, "")
	}

	err = f.SaveAs(dst + "/" + filename.String() + ".xlsx")
	if err != nil {
		return nil, err
	}
	_, err = minioClient.FPutObject(bucketName, filename.String()+".xlsx", dst+"/"+filename.String()+".xlsx", minio.PutObjectOptions{ContentType: excelContentType})
	if err != nil {
		return nil, err
	}
	os.Remove(dst + "/" + filename.String() + ".xlsx")

	return &Response{
		FileName: filename.String(),
		Url:      "https://" + minioEndpoint + "/" + bucketName + "/" + filename.String() + ".xlsx",
	}, nil
}
