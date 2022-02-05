package excel_generator

import (
	"encoding/json"
	"fmt"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/xtgo/uuid"
)

func GenerateExcel(data []byte, bucketName, minioEndpoint, accessKey, secretKey string, titles ...Pair) (*Response, error) {
	var model Data

	err := json.Unmarshal(data, &model.Rows)
	if err != nil {
		return nil, err
	}

	f := excelize.NewFile()
	filename := uuid.NewRandom()

	column := 'A'
	for _, v := range titles {
		f.SetCellValue("Sheet1", fmt.Sprintf("%c%d", column, 1), v.Second)
		column++
	}

	style, _ := f.NewStyle(`{"alignment":{"horizontal":"center", "vertical": "center", "wrap_text": true}}`)
	f.SetCellStyle("Sheet1", "A1", fmt.Sprintf("%c%d", column, len(model.Rows)+1), style)
	f.SetColWidth("Sheet1", "B", string(column), 30)
	f.SetRowHeight("Sheet1", 1, 35)

	writer(model, f, titles)

	v, err := minioUploader(bucketName, minioEndpoint, accessKey, secretKey, f, filename)
	if err != nil {
		return nil, err
	}

	return v, nil
}

// GenerateWithWorkers generate excel file with worker pool, return FileName, channel for sending job, channel for receive percent, response channel
func GenerateWithWorkers(workerCount int, jobsCount int, bucketName, minioEndpoint, accessKey, secretKey string, titles ...Pair) (string, chan<- []byte, <-chan int, <-chan Result) {
	filename := uuid.NewRandom()
	f := excelize.NewFile()

	// writing title documents
	column := 'A'
	for _, v := range titles {
		f.SetCellValue("Sheet1", fmt.Sprintf("%c%d", column, 1), v.Second)
		column++
	}

	// styling documents
	style, _ := f.NewStyle(`{"alignment":{"horizontal":"center", "vertical": "center", "wrap_text": true}}`)
	f.SetCellStyle("Sheet1", "A1", fmt.Sprintf("%c%d", column, jobsCount+1), style)
	f.SetColWidth("Sheet1", "B", string(column), 35)
	f.SetRowHeight("Sheet1", 1, 35)

	var percent int
	jobs := make(chan WorkerJob, jobsCount)
	results := make(chan error, jobsCount)
	percentChan := make(chan int, 101)
	jobChan := make(chan []byte, jobsCount)
	responseChan := make(chan Result, 1)

	// running workers
	for k := 1; k <= workerCount; k++ {
		go workers(jobs, results, f, &percent, jobsCount, percentChan, titles)
	}

	// receiving jobs
	go sendJobs(
		jobs,
		results,
		jobChan,
		jobsCount,
		responseChan,
		filename,
		f,
		MinioConfigurations{
			BucketName:    bucketName,
			MinioEndpoint: minioEndpoint,
			AccessKey:     accessKey,
			SecretKey:     secretKey,
		},
	)

	return filename.String(), jobChan, percentChan, responseChan
}
