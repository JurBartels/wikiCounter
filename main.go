package main

import (
  "fmt"
  "net/http"
  "encoding/json"
  "strconv"
)

func main(){

  for i := 1; i <= 20; i++ {
    if i != 4 {
      char_count, err := getCharCount(i)
      if err != nil {
        fmt.Println(err)
      }
      fmt.Println(i)
      fmt.Println(char_count)
    }
  }
}

func getCharCount(grpnum int) (float64, error) {
  conc := strconv.Itoa(grpnum)
  var f interface{}
  err := getJSON("http://cstwiki.wtb.tue.nl/api.php?action=query&prop=info&titles=PRE2016_3_Groep"+ conc +"&format=json", &f)
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
