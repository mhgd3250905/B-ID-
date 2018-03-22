package main

import (
	"encoding/json"
	"fmt"
	"github.com/PuerkitoBio/goquery"
	"log"

	"github.com/tealeg/xlsx"
	"os"
	"strings"
	"time"

)


type BiliFollowersBean struct {
	Data BiliData
}

type BiliData struct {
	Follower int
}

type WeiboUserInfo struct {
	Followers_count int
}

type WeiboData struct {
	UserInfo WeiboUserInfo
}

type WeiboFollowersBean struct {
	Data WeiboData
}

//哔哩哔哩的获取粉丝数的URL
var biliUrl = "https://api.bilibili.com/x/relation/stat?jsonp=jsonp&vmid=%s"

//微博获取粉丝数的URL
var weiboUrl = "https://m.weibo.cn/api/container/getIndex?type=uid&value=%s"

var biliIdIndex = 2

var weiboIdIndex = 3

var excelFile xlsx.File

func main() {
	fmt.Printf("亲爱的小白!欢迎进入开哥帮助环节！")

	fmt.Println("即将开始释放爬虫...")

	biliUpMap := runBiliSpider()
	fmt.Println("返回成功")

	weiboUpMap := runWeiboSpider()
	fmt.Println("返回成功")

	fmt.Println("开始保存数据...")
	err, fileName := saveExcel(biliUpMap, weiboUpMap)
	if err != nil {
		fmt.Println("保存数据失败!")
	} else {
		fmt.Println()
		fmt.Printf("保存数据成功，文件名为%s。", fileName)
		fmt.Println()
	}

	fmt.Printf("亲爱的小白!欢迎下次光临！输入88+回车可关闭此窗口！")

	fmt.Println()
	var exit int
	fmt.Scan(&exit)
	if exit==88 {

	}else {
		fmt.Println("说了输入88，都不说再见的吗？")
		fmt.Scan(&exit)
	}
}

func runBiliSpider() map[string]BiliFollowersBean {
	fmt.Println("爬虫已经进入bilibili，正在获取数据...")
	ids := GetData(biliIdIndex)
	// 先声明map
	var upMap map[string]BiliFollowersBean
	// 再使用make函数创建一个非nil的map，nil map不能赋值
	upMap = make(map[string]BiliFollowersBean)
	for _, id := range ids {
		//获取完整的URL
		totlaUrl := fmt.Sprintf(biliUrl, id)
		//获取数据
		doc, err := goquery.NewDocument(totlaUrl)
		if err != nil {
			log.Fatal(err)
		}
		//获取结果
		jsonStr := doc.Text()
		//转化为结构体
		item := biliStr2Json(jsonStr)

		//fmt.Println(item)
		//加入到字典
		upMap[id] = item
	}
	fmt.Println("获取完毕，正在装载返回。")
	//返回数据类
	return upMap
}

func runWeiboSpider() map[string]WeiboFollowersBean {
	fmt.Println("爬虫已经进入微博，正在获取数据...")
	ids := GetData(weiboIdIndex)
	// 先声明map
	var upMap map[string]WeiboFollowersBean
	// 再使用make函数创建一个非nil的map，nil map不能赋值
	upMap = make(map[string]WeiboFollowersBean)
	for _, id := range ids {
		//获取完整的URL
		totlaUrl := fmt.Sprintf(weiboUrl, id)
		//fmt.Println(totlaUrl)
		//获取数据
		doc, err := goquery.NewDocument(totlaUrl)
		if err != nil {
			log.Fatal(err)
		}
		//获取结果
		jsonStr := doc.Text()
		//fmt.Println(jsonStr)
		var item WeiboFollowersBean
		//转化为结构体
		json.Unmarshal([]byte(jsonStr), &item)

		//fmt.Println(item)
		//加入到字典
		upMap[id] = item
	}
	fmt.Println("获取完毕，正在装载返回。")
	return upMap
}

//保存文件
func saveExcel(map1 map[string]BiliFollowersBean, map2 map[string]WeiboFollowersBean) (error, string) {
	//获取Excel
	xlFile, err := xlsx.OpenFile("ids.xlsx")
	if err != nil {
		fmt.Println("打开Excel失败")
	}

	for _, sheet := range xlFile.Sheets {
		//获取名字Sheet1的sheet
		if strings.EqualFold(sheet.Name, "Sheet1") {
			for rowIndex, row := range sheet.Rows {
				if rowIndex != 0 {
					follows := map1[row.Cells[2].Value]
					followerCount := follows.Data.Follower
					cell := row.AddCell()
					cell.SetValue(followerCount)
				}
			}
		}
	}
	for _, sheet := range xlFile.Sheets {
		//获取名字Sheet1的sheet
		if strings.EqualFold(sheet.Name, "Sheet1") {
			for rowIndex, row := range sheet.Rows {
				if rowIndex != 0 {
					follows := map2[row.Cells[3].Value]
					followerCount := follows.Data.UserInfo.Followers_count
					cell := row.AddCell()
					cell.SetValue(followerCount)
				}
			}
		}
	}

	//保存
	saveFilePath := fmt.Sprintf("followrs_%s.xlsx", time.Now().Format("2006_01_02_15_04_05"))

	isExist, err := PathExists(saveFilePath)
	if err != nil {
		fmt.Println("获取文件是否存在失败！")
		return err, saveFilePath
	}

	if isExist {
		err := os.Remove(saveFilePath) //删除文件test.txt
		if err != nil {
			//如果删除失败则输出 file remove Error!
			fmt.Println("file remove Error!")
			//输出错误详细信息
			fmt.Printf("%s", err)
			return err, saveFilePath
		} else {
			//如果删除成功则输出 file remove OK!
			fmt.Print("file remove OK!")
		}
	}
	err = xlFile.Save(saveFilePath)
	if err != nil {
		fmt.Println("保存失败")
		return err, saveFilePath
	}

	return nil, saveFilePath
}

//保存Bilibili的Excel
func saveBiliExcel(upMap map[string]BiliFollowersBean) {
	fmt.Println("开始进行保存Excel数据----->")
	//获取Excel
	xlFile, err := xlsx.OpenFile("id.xlsx")
	if err != nil {
		fmt.Println("打开Excel失败")
	}

	for _, sheet := range xlFile.Sheets {
		//获取名字Sheet1的sheet
		if strings.EqualFold(sheet.Name, "Sheet1") {
			for rowIndex, row := range sheet.Rows {
				if rowIndex != 0 {
					follows := upMap[row.Cells[2].Value]
					followerCount := follows.Data.Follower
					cell := row.AddCell()
					cell.SetValue(followerCount)
				}
			}
		}
	}

	saveFilePath := fmt.Sprintf("followrs_%s.xlsx", time.Now().Format("2006_01_02_15_04_05"))

	isExist, err := PathExists(saveFilePath)
	if err != nil {
		fmt.Println("获取文件是否存在失败！")
	}

	if isExist {
		err := os.Remove(saveFilePath) //删除文件test.txt
		if err != nil {
			//如果删除失败则输出 file remove Error!
			fmt.Println("file remove Error!")
			//输出错误详细信息
			fmt.Printf("%s", err)
		} else {
			//如果删除成功则输出 file remove OK!
			fmt.Print("file remove OK!")
		}
	}
	err = xlFile.Save(saveFilePath)
	if err != nil {
		fmt.Println("保存失败")
	}
}

//保存Weibo的Excel
func saveWeiboExcel(upMap map[string]WeiboFollowersBean) {
	fmt.Println("开始进行保存Excel数据----->")
	//获取Excel
	xlFile, err := xlsx.OpenFile("ids.xlsx")
	if err != nil {
		fmt.Println("打开Excel失败")
	}

	for _, sheet := range xlFile.Sheets {
		//获取名字Sheet1的sheet
		if strings.EqualFold(sheet.Name, "Sheet1") {
			for rowIndex, row := range sheet.Rows {
				if rowIndex != 0 {
					follows := upMap[row.Cells[2].Value]
					followerCount := follows.Data.UserInfo.Followers_count
					cell := row.AddCell()
					cell.SetValue(followerCount)
				}
			}
		}
	}

	saveFilePath := fmt.Sprintf("followrs_%s.xlsx", time.Now().Format("2006_01_02_15_04_05"))

	isExist, err := PathExists(saveFilePath)
	if err != nil {
		fmt.Println("获取文件是否存在失败！")
	}

	if isExist {
		err := os.Remove(saveFilePath) //删除文件test.txt
		if err != nil {
			//如果删除失败则输出 file remove Error!
			fmt.Println("file remove Error!")
			//输出错误详细信息
			fmt.Printf("%s", err)
		} else {
			//如果删除成功则输出 file remove OK!
			fmt.Print("file remove OK!")
		}
	}
	err = xlFile.Save(saveFilePath)
	if err != nil {
		fmt.Println("保存失败")
	}
}

//哔哩哔哩String 2 json
func biliStr2Json(jsonStr string) (biliBean BiliFollowersBean) {
	json.Unmarshal([]byte(jsonStr), &biliBean)
	return biliBean
}

/**
获取Excel中的Ids
*/
func GetData(idIndex int) []string {
	//fmt.Println("开始进行获取Excel数据----->")

	xlFile, err := xlsx.OpenFile("/ids.xlsx")
	if err != nil {
		fmt.Println("打开Excel失败,",err)
	}
	var ids []string
	for _, sheet := range xlFile.Sheets {
		//获取名字Sheet1的sheet
		if strings.EqualFold(sheet.Name, "Sheet1") {
			for rowIndex, row := range sheet.Rows {
				if rowIndex != 0 {
					for cellIndex, cell := range row.Cells {
						if cellIndex == idIndex {
							ids = append(ids, cell.Value)
						}
					}
				}
			}
		}
	}
	//fmt.Println(ids)
	return ids
}

func PathExists(path string) (bool, error) {
	_, err := os.Stat(path)
	if err == nil {
		return true, nil
	}
	if os.IsNotExist(err) {
		return false, nil
	}
	return false, err
}
