package excel_generator

import (
	"encoding/json"
	"fmt"
	"os"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/minio/minio-go"
)

func getExcelColumnChar(num int) string {
	var result string
	for num > 0 {
		result = string('A'+rune((num-1)%26)) + result
		num = (num - 1) / 26
	}
	return result
}

// writer write all data on excel file
func writer(data Data, f *excelize.File, titles []Pair) {
	key := 1
	for _, row := range data.Rows {
		key++
		column := 1
		f.SetCellValue("Sheet1", fmt.Sprintf("%c%d", 'A', key), key-1)
		for index := range titles {
			if value, ok := row[titles[index].First]; ok {
				column++
				f.SetCellValue("Sheet1", fmt.Sprintf("%s%d", getExcelColumnChar(column), key), value)
			}
		}
	}
}

// minioUploader upload created file to the minio
func minioUploader(bucketName, minioEndpoint, accessKey, secretKey string, f *excelize.File, filename string) (*Response, error) {
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

	err = f.SaveAs(dst + "/" + filename + ".xlsx")
	if err != nil {
		return nil, err
	}
	_, err = minioClient.FPutObject(bucketName, filename+".xlsx", dst+"/"+filename+".xlsx", minio.PutObjectOptions{ContentType: excelContentType})
	if err != nil {
		return nil, err
	}

	os.Remove(dst + "/" + filename + ".xlsx")

	return &Response{
		FileName: filename,
		Url:      "https://" + minioEndpoint + "/" + bucketName + "/" + filename + ".xlsx",
	}, nil
}

func sendJobs(jobs chan WorkerJob, results chan error, job chan []byte, jobsCount int, response chan<- Result, filename string, f *excelize.File, config MinioConfigurations) {
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
	for job := range jobs {
		if ((job.Row-1)*100)/jobsCount > *percent {
			*percent = ((job.Row - 1) * 100) / jobsCount
			percentChan <- *percent
		}

		column := 1
		f.SetCellValue("Sheet1", fmt.Sprintf("%s%d", getExcelColumnChar(column), job.Row), job.Row-1)
		for index := range titles {
			if value, ok := job.Data[titles[index].First]; ok {
				column++
				f.SetCellValue("Sheet1", fmt.Sprintf("%s%d", getExcelColumnChar(column), job.Row), value)
			}
		}

		results <- nil
	}
}
