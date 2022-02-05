package excel_generator

type Data struct {
	Rows []map[string]interface{} `json:"rows"`
}

type Response struct {
	FileName string `json:"file_name"`
	Url      string `json:"url"`
}

type Pair struct {
	First  string `json:"first"`
	Second string `json:"second"`
}

type Result struct {
	Response *Response `json:"response"`
	Error    error     `json:"error"`
}

type WorkerJob struct {
	Row  int                    `json:"row"`
	Data map[string]interface{} `json:"data"`
}

type MinioConfigurations struct {
	BucketName    string `json:"bucket_name"`
	MinioEndpoint string `json:"minio_endpoint"`
	AccessKey     string `json:"access_key"`
	SecretKey     string `json:"secret_key"`
}

func NewPair(first, second string) Pair {
	return Pair{
		First:  first,
		Second: second,
	}
}
