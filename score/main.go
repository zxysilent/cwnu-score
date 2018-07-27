// score project main.go
package main

import (
	"fmt"
	"io/ioutil"
	"net/http"
	"net/http/cookiejar"
	"net/url"
	"strconv"
	"strings"
	"time"

	"github.com/astaxie/beego/config"
	"github.com/djimenez/iconv-go"
	"github.com/tealeg/xlsx"
)

const (
	HomeUrl  = "http://210.41.193.28/web/web/mis/"
	LoginUrl = "http://210.41.193.28/jiaoshi/bangong/main/check.asp"
	ScoreUrl = "http://210.41.193.28/jiaoshi/xslm/cj/note"
)

func main() {
	defer func() {
		if err := recover(); err != nil {
			fmt.Println("出错:", err)
			time.Sleep(time.Second * 5)
		}
	}()
	config, err := config.NewConfig("ini", "app.conf")
	if err != nil {
		panic(err)
	}
	num := config.String("num")
	pass := config.String("pass")
	fmt.Println("学号:", num, "密码:", pass)
	scoreHtml := getScoreHtml(num, pass)
	if !strings.Contains(scoreHtml, "成绩查询") {
		panic("密码错误或者未评教完")
	}
	scoreHtml = GetBetweenStr(scoreHtml, "<tr height=25", `</table>`)
	scoreHtml = strings.Replace(scoreHtml, "class=g_body_2", "", -1)
	scoreHtml = strings.Replace(scoreHtml, "height=25 ", "", -1)
	scoreHtml = strings.Replace(scoreHtml, "class=g_body_1", "", -1)
	scoreHtml = strings.Replace(scoreHtml, "align=center", "", -1)
	scoreHtml = strings.Replace(scoreHtml, " ", "", -1)
	scoreHtml = strings.Replace(scoreHtml, "\r\n", "", -1)
	scoreHtml = strings.TrimPrefix(scoreHtml, "<tr>")
	scoreHtml = strings.TrimSuffix(scoreHtml, "</tr>")
	trs := strings.Split(scoreHtml, "</tr><tr>")
	//	slicesResult := []Result{}
	//写入 excel
	var file *xlsx.File
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var cell *xlsx.Cell

	file = xlsx.NewFile()
	sheet, err = file.AddSheet(num + "成绩")
	if err != nil {
		fmt.Printf(err.Error())
	}
	row = sheet.AddRow()
	cell = row.AddCell()
	cell.Value = "编号"
	cell = row.AddCell()
	cell.Value = "学期"
	cell = row.AddCell()
	cell.Value = "课程"
	cell = row.AddCell()
	cell.Value = "课程性质"
	cell = row.AddCell()
	cell.Value = "学分"
	cell = row.AddCell()
	cell.Value = "成绩"
	cell = row.AddCell()
	cell.Value = "绩点"
	cell = row.AddCell()
	cell.Value = "学分绩点"
	cell = row.AddCell()
	cell.Value = "备注"
	cell = row.AddCell()
	cell.Value = "修读方式"
	for i := 0; i < len(trs); i++ {
		tr := strings.TrimPrefix(trs[i], "<td>")
		tr = strings.TrimSuffix(tr, "</td>")
		tds := strings.Split(tr, "</td><td>")
		row = sheet.AddRow()
		cell = row.AddCell()
		cell.Value = strconv.Itoa(i + 1)
		cell = row.AddCell()
		cell.Value = tds[0]
		cell = row.AddCell()
		cell.Value = tds[1]
		cell = row.AddCell()
		cell.Value = tds[2]
		cell = row.AddCell()
		cell.Value = tds[3]
		cell = row.AddCell()
		cell.Value = tds[4]
		cell = row.AddCell()
		cell.Value = tds[5]
		cell = row.AddCell()
		cell.Value = tds[6]
		cell = row.AddCell()
		cell.Value = tds[7]
		cell = row.AddCell()
		cell.Value = tds[8]
	}
	err = file.Save("成绩" + time.Now().Format("2006-01-02 150405") + ".xlsx")
	if err != nil {
		fmt.Printf(err.Error())
	}
	fmt.Println("v1 over ")
}
func GetBetweenStr(str, start, end string) string {
	n := strings.Index(str, start)
	if n == -1 {
		n = 0
	}
	str = string([]byte(str)[n:])
	m := strings.Index(str, end)
	if m == -1 {
		m = len(str)
	}
	str = string([]byte(str)[:m])
	return str
}

// getScoreHtml
func getScoreHtml(num string, pass string) string {
	// HTTP client.
	var client http.Client
	// Cookie jar.
	jar, err := cookiejar.New(nil)
	if err != nil {
		panic(err)
	}
	client.Jar = jar
	// Init cookie
	if resp, err := client.Get(HomeUrl); err != nil {
		panic(err)
	} else {
		resp.Body.Close()
	}

	// login begin
	v := url.Values{}
	v.Set("user", num)
	v.Set("pwd", "")
	v.Set("user1", encrypt(num))
	v.Set("pwd1", encrypt(pass))
	bodyLogin := ioutil.NopCloser(strings.NewReader(v.Encode()))
	reqestLogin, err := http.NewRequest("POST", LoginUrl, bodyLogin)
	if err != nil {
		panic(err)
	}
	addHeader(reqestLogin)
	if resp, err := client.Do(reqestLogin); err != nil {
		panic(err)
	} else {
		resp.Body.Close()
	}

	// login end

	// score
	v1 := url.Values{}
	v1.Set("lanmuys", "学生栏目")
	bodyScore := ioutil.NopCloser(strings.NewReader(v1.Encode()))
	reqestScore, err := http.NewRequest("POST", ScoreUrl, bodyScore)
	if err != nil {
		panic(err)
	}
	addHeader(reqestScore)
	resp, err := client.Do(reqestScore)
	if err != nil {
		panic(err)
	}
	defer resp.Body.Close()
	body, _ := ioutil.ReadAll(resp.Body)
	out := make([]byte, len(body))
	iconv.Convert(body, out, "gb2312", "utf-8")
	return string(out)
}

// addHeader
func addHeader(reqest *http.Request) {
	reqest.Header.Add("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8")
	reqest.Header.Add("Accept-Encoding", "gzip, deflate")
	reqest.Header.Add("Accept-Language", "zh-CN,zh;q=0.8")
	reqest.Header.Add("Cache-Control", "no-cache")
	reqest.Header.Add("Content-Length", "143")
	reqest.Header.Add("Content-Type", "application/x-www-form-urlencoded")
	reqest.Header.Add("Host", "218.6.128.130")
	reqest.Header.Add("Origin", "http://218.6.128.130")
	reqest.Header.Add("Pragma", "no-cache")
	reqest.Header.Add("Proxy-Connection", "keep-alive")
	reqest.Header.Add("Referer", "http://218.6.128.130/web/web/mis/")
	reqest.Header.Add("Upgrade-Insecure-Requests", "1")
	reqest.Header.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0.14393; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2950.5 Safari/537.36")
}

//encrypt user password and uers num
func encrypt(str string) string {
	var strs []string
	for i := 0; i < len(str); i++ {
		strs = append(strs, strconv.Itoa(int(str[i])+100000))
	}
	return strings.Join(strs, "")
}
