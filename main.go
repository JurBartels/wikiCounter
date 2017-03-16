package main

import (
  "fmt"
  "net/http"
  "encoding/json"
  "strconv"
  "github.com/tealeg/xlsx"
  "os"
  "time"
)

var filename = "log.xlsx"

func main(){

  var file *xlsx.File
  var err error

  if  _, err = os.Stat(filename); os.IsNotExist(err){
    fmt.Println("file not found")
    file , err = createFile(filename)
    if err != nil {
      fmt.Println(err)
    }
  } else {
    file,err = xlsx.OpenFile(filename)
    if err != nil {
      fmt.Println(err)
    }
  }

  results := make(map[int]float64)

  for i := 1; i <= 20; i++ {
    if i != 4 {
      char_count, err := getCharCount(i)
      if err != nil {
        fmt.Println(err)
      }
      results[i] = char_count
    }
  }

  key, er := readKey(file)
  if er != nil{
    fmt.Println(er)
  }
  updateKey(filename, file, key+1)
  writeToExel(int(key), filename, results, file)
}

func getCharCount(grpnum int) (float64, error) {
  conc := strconv.Itoa(grpnum)
  var f interface{}
  err := getJSON("http://cstwiki.wtb.tue.nl/api.php?action=query&prop=info&" +
    "titles=PRE2016_3_Groep"+ conc +"&format=json", &f)
  if err != nil{
    return 0,err
  }
  msg := f.(map[string]interface{})
  decoded_msg := msg["query"].
    (map[string]interface{})["pages"].
    (map[string]interface{})

  var key string
  for k := range decoded_msg{
    key = k
  }
  return decoded_msg[key].(map[string]interface{})["length"].(float64), nil
}

func getJSON(url string, target interface{}) error {
  r, err := http.Get(url)
  if err != nil {
    return err
  }
  defer r.Body.Close()
  return json.NewDecoder(r.Body).Decode(target)
}

func writeToExel(col int, filename string, resultmap map[int]float64, file *xlsx.File) error {
  sheet := file.Sheets[0]
  t := time.Now().Format("02 Jan 06 15:04 MST")
  cell := sheet.Cell(0,col)
  cell.Value = t
  for i,e := range resultmap {
    if i != 4 {
      cell := sheet.Cell(i,col)
      cell.SetFloat(e)
    }
  }
  err := file.Save(filename)
  if err != nil{
    return err
  }
  return nil
}

func createFile(filename string) (*xlsx.File , error) {
  file := xlsx.NewFile()
  sheet, err := file.AddSheet("Sheet1")
  if err != nil {
      return nil,err
  }
  fillNewFile(sheet)
  err = file.Save(filename)
  if err != nil {
      return nil,err
  }
  return file, nil
}

func fillNewFile(sheet *xlsx.Sheet) {
  for i := 1;i <= 20 ; i++ {
    cell := sheet.Cell(i,0)
    cell.Value  = "Group " + strconv.Itoa(i)
  }
  cell := sheet.Cell(22,0)
  cell.SetFloat(1)
}

func readKey(file *xlsx.File) (float64, error) {
  sheet := file.Sheets[0]
  cell := sheet.Cell(22,0)
  current_key, err := cell.Float()
  return current_key,err
}

func updateKey(filename string, file *xlsx.File, val float64){
  sheet := file.Sheets[0]
  cell := sheet.Cell(22,0)
  cell.SetFloat(val)
  file.Save(filename)
}
