package excel_generator

import (
	"encoding/json"
	"errors"
	"fmt"
	"os"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/minio/minio-go"
	"github.com/xtgo/uuid"
)

// writer write all data on excel file
func writer(data Data, f *excelize.File, titles []Pair) {
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

// minioUploader uploads the created file to minio
func minioUploader(bucketName, minioEndpoint, accessKey, secretKey string, f *excelize.File, filename uuid.UUID) (*Response, error) {
	dst, _ := os.Getwd()

	excelContentType := "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

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

func sendJobs(jobs chan WorkerJob, results chan error, job chan []byte, jobsCount int, response chan<- Result, filename uuid.UUID, f *excelize.File, config MinioConfigurations) {
	for i := 0; i < jobsCount; i++ {
		var data WorkerJob

		err := json.Unmarshal(<-job, &data.Data)
		if err != nil {
			response <- Result{
				Response: nil,
				Error:    err,
			}
			return
		}

		data.Row = i + 2

		// doing job
		jobs <- data
	}

	close(jobs)

	for i := 0; i < jobsCount; i++ {
		err := <-results
		if err != nil {
			response <- Result{
				Response: nil,
				Error:    err,
			}
			return
		}
	}

	close(results)

	resp, err := minioUploader(config.BucketName, config.MinioEndpoint, config.AccessKey, config.SecretKey, f, filename)
	response <- Result{
		Response: resp,
		Error:    err,
	}
}

// workers
func workers(jobs <-chan WorkerJob, results chan<- error, f *excelize.File, percent *int, jobsCount int, percentChan chan<- int, titles []Pair) {
	if len(titles) > 25 {
		results <- errors.New("rows limit error, you can generate max 25 rows")
		return
	}

	for job := range jobs {
		if ((job.Row-1)*100)/jobsCount > *percent {
			*percent = ((job.Row - 1) * 100) / jobsCount
			percentChan <- *percent
		}

		column := 'A'
		f.SetCellValue("Sheet1", fmt.Sprintf("%c%d", column, job.Row), job.Row-1)
		for _, v := range titles {
			for k, value := range job.Data {
				if v.First == k {
					column++
					f.SetCellValue("Sheet1", fmt.Sprintf("%c%d", column, job.Row), value)
					break
				}
			}
		}

		results <- nil
	}
}
