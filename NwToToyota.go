package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"
	"unicode/utf8"

	"github.com/tealeg/xlsx"
	"golang.org/x/text/encoding/japanese"
	"golang.org/x/text/transform"
	"golang.org/x/text/unicode/norm"
)

func failOnError(err error) {
	if err != nil {
		log.Fatal("Error:", err)
	}
}

func main() {
	flag.Parse()

	// ログファイル準備
	logfile, err := os.OpenFile("./log.txt", os.O_APPEND|os.O_CREATE|os.O_WRONLY, os.ModePerm)
	failOnError(err)
	defer logfile.Close()

	log.SetOutput(logfile)

	log.Print("Start\r\n")

	// ファイルを読み込んで二次元配列に入れる
	records := readfile(flag.Arg(0))

	// 出力する会社を調査
	coRecods := coSurvey(records)

	//出力するフォルダを作成
	outDir := dirCreate(flag.Arg(0))

	// データの変換
	dataConversion(outDir, records, coRecods)

	// 受診者名簿の作成
	meiboCreate(outDir, records, coRecods)

	log.Print("Finish !\r\n")

}

func readfile(filename string) [][]string {
	//入力ファイル準備
	infile, err := os.Open(filename)
	failOnError(err)
	defer infile.Close()

	reader := csv.NewReader(transform.NewReader(infile, japanese.ShiftJIS.NewDecoder()))
	reader.Comma = '\t'

	//CSVファイルを２次元配列に展開
	readrecords := make([][]string, 0)
	for {
		record, err := reader.Read() // 1行読み出す
		if err == io.EOF {
			break
		} else {
			failOnError(err)
		}

		readrecords = append(readrecords, record)
	}

	return readrecords
}

func coSurvey(records [][]string) [][]string {
	companys := [][]string{{"2000100100000001", "トヨタモビリティ東京（株）", "0"},
		{"2000100100000026", "ティーシーサービス（株）", "0"},
		{"9500100100000001", "トヨタモビリティ東京（株）", "0"},
		{"2000100100000020", "ＴＭプロサービス（株）", "0"},
		{"2000100100000008", "（株）トヨテック", "0"},
		{"2000100100009002", "トヨタ東京カローラ（株）", "0"},
		{"2000100100009004", "（株）センチュリーサービス", "0"},
	}

	coRecMax := len(records)
	for i := 1; i < coRecMax; i++ {
		for _, com := range companys {
			if com[0] == records[i][4] {
				count, _ := strconv.Atoi(com[2])
				com[2] = fmt.Sprint(count + 1)
				break
			}
		}
	}

	outCompanys := make([][]string, 0)
	for j, _ := range companys {
		if companys[j][2] != "0" {
			outCompanys = append(outCompanys, companys[j])
			// log.Print(companys[j][2] + " " + companys[j][0] + ":" + companys[j][1] + "\r\n")
		}
	}

	return outCompanys

}

func dirCreate(path string) string {
	day := time.Now()
	outDir, _ := filepath.Split(path)
	outDirPlus := outDir + "/トヨタモビリティ東京" + day.Format("20060102")

	if err := os.Mkdir(outDirPlus, 0777); err != nil {
		log.Print(outDirPlus + "\r\n")
		log.Print("出力先のディレクトリを作成できませんでした\r\n")
		return outDir
	} else {
		return outDirPlus + "/"
	}
}

func dataConversion(filename string, inRecs [][]string, coRecs [][]string) {
	// var excelFile *xlsx.File
	// var sheet *xlsx.Sheet
	var vcell *xlsx.Cell
	var r int
	var cell string

	recLen := 222 //出力するレコードの項目数
	cRec := make([]string, recLen)
	var I int

	day := time.Now()

	//会社毎に健診データファイルを作成する
	for _, coRec := range coRecs {

		/*
			outfile, err := os.Create(filename + coRec[1] + "健診データ" + day.Format("20060102") + ".txt")
			failOnError(err)
			defer outfile.Close()

			writer := csv.NewWriter(transform.NewWriter(outfile, japanese.ShiftJIS.NewEncoder()))
			writer.Comma = '\t'
			writer.UseCRLF = true
		*/

		excelName := filename + coRec[1] + "健診データ" + day.Format("20060102") + ".xlsx"
		excelFile := xlsx.NewFile()
		xlsx.SetDefaultFont(11, "ＭＳ Ｐゴシック")
		sheet, err := excelFile.AddSheet("データ")
		failOnError(err)

		// 1行目（タイトル）
		for I, _ = range cRec {
			cRec[I] = ""
		}

		cRec[0] = "idou.sya_bg"
		cRec[19] = "knk_kenkork.jushin_date"
		cRec[22] = "knk_kenkork_kensa.kensa_val_071"
		cRec[27] = "knk_kenkork_kensa.kensa_val_005"
		cRec[28] = "knk_kenkork_kensa.kensa_val_006"
		cRec[29] = "knk_kenkork_kensa.kensa_val_007"
		cRec[31] = "knk_kenkork_kensa.kensa_val_008"
		cRec[33] = "knk_kenkork_kensa.kensa_val_001"
		cRec[34] = "knk_kenkork_kensa.kensa_val_002"
		cRec[35] = "knk_kenkork_kensa.kensa_val_003"
		cRec[36] = "knk_kenkork_kensa.kensa_val_023"
		cRec[37] = "knk_kenkork_kensa.kensa_val_027"
		cRec[38] = "knk_kenkork_kensa.kensa_val_025"
		cRec[39] = "knk_kenkork_kensa.kensa_val_024"
		cRec[40] = "knk_kenkork_kensa.kensa_val_028"
		cRec[41] = "knk_kenkork_kensa.kensa_val_026"
		cRec[44] = "knk_kenkork_kensa.kensa_val_039"
		cRec[45] = "knk_kenkork_kensa.kensa_val_038"
		cRec[46] = "knk_kenkork_kensa.kensa_val_037"
		cRec[48] = "knk_kenkork_kensa.kensa_val_033"
		cRec[49] = "knk_kenkork_kensa.kensa_val_034"
		cRec[50] = "knk_kenkork_kensa.kensa_val_035"
		cRec[51] = "knk_kenkork_Kensa.kensa_val_080"
		cRec[54] = "knk_kenkork_kensa.kensa_val_041"
		cRec[55] = "knk_kenkork_kensa.kensa_val_079"
		cRec[57] = "knk_kenkork_kensa.kensa_val_042"
		cRec[64] = "knk_kenkork_kensa.kensa_val_031"
		cRec[65] = "knk_kenkork_kensa.kensa_val_030"
		cRec[70] = "knk_kenkork_kensa.kensa_val_047"
		cRec[72] = "knk_kenkork_kensa.kensa_val_021"
		cRec[78] = "knk_kenkork_kensa.kensa_val_010"
		cRec[79] = "knk_kenkork_kensa.kensa_val_011"
		cRec[80] = "knk_kenkork_kensa.kensa_val_012"
		cRec[81] = "knk_kenkork_kensa.kensa_val_013"
		cRec[168] = "knk_kenkork_kensa.kensa_val_072"
		cRec[176] = "knk_kenkork_kensa.kensa_val_049"
		cRec[179] = "knk_kenkork_kensa.kensa_val_050"
		cRec[182] = "knk_kenkork_kensa.kensa_val_051"
		cRec[185] = "knk_kenkork_kensa.kensa_val_052"
		cRec[186] = "knk_kenkork_kensa.kensa_val_053"
		cRec[187] = "knk_kenkork_kensa.kensa_val_054"
		cRec[188] = "knk_kenkork_kensa.kensa_val_055"
		cRec[189] = "knk_kenkork_kensa.kensa_val_056"
		cRec[190] = "knk_kenkork_kensa.kensa_val_057"
		cRec[191] = "knk_kenkork_kensa.kensa_val_058"
		cRec[192] = "knk_kenkork_kensa.kensa_val_059"
		cRec[193] = "knk_kenkork_kensa.kensa_val_060"
		cRec[194] = "knk_kenkork_kensa.kensa_val_061"
		cRec[195] = "knk_kenkork_kensa.kensa_val_081"
		cRec[196] = "knk_kenkork_kensa.kensa_val_062"
		cRec[197] = "knk_kenkork_kensa.kensa_val_063"
		cRec[198] = "knk_kenkork_kensa.kensa_val_064"
		cRec[199] = "knk_kenkork_kensa.kensa_val_082"
		cRec[200] = "knk_kenkork_kensa.kensa_val_065"
		cRec[201] = "knk_kenkork_kensa.kensa_val_066"
		cRec[202] = "knk_kenkork_kensa.kensa_val_067"
		cRec[203] = "knk_kenkork_kensa.kensa_val_068"
		cRec[204] = "knk_kenkork_kensa.kensa_val_069"
		cRec[205] = "knk_kenkork_kensa.kensa_val_070"
		cRec[212] = "knk_kenkork_kensa.kensa_val_020"
		cRec[213] = "knk_kenkork_kensa.hantei_val_020"
		cRec[214] = "knk_kenkork_kensa.kensa_val_044"
		cRec[215] = "knk_kenkork_kensa.kensa_val_045"
		cRec[216] = "knk_kenkork_kensa.kensa_val_016"
		cRec[217] = "knk_kenkork_kensa.kensa_val_017"
		cRec[218] = "knk_kenkork_kensa.kensa_val_018"
		cRec[219] = "knk_kenkork_kensa.kensa_val_019"
		cRec[220] = "knk_kenkork_kensa.kensa_val_046"
		cRec[221] = "knk_kenkork_kensa.hantei_val_046"
		//writer.Write(cRec)
		row := sheet.AddRow()
		for _, cell = range cRec {
			//sheet.Cell(0, c).Value = cell
			vcell = row.AddCell()
			vcell.Value = cell
		}

		// 2行目（タイトル）
		for I, _ = range cRec {
			cRec[I] = ""
		}

		cRec[0] = "#社員番号"
		cRec[19] = "受診日付"
		cRec[22] = "検査コード071_医療機関側判定結果"
		cRec[27] = "検査コード005_医療機関側検査値"
		cRec[28] = "検査コード006_医療機関側検査値"
		cRec[29] = "検査コード007_医療機関側検査値"
		cRec[31] = "検査コード008_医療機関側検査値"
		cRec[33] = "検査コード001_医療機関側検査値"
		cRec[34] = "検査コード02_医療機関側判定結果"
		cRec[35] = "検査コード003_医療機関側検査値"
		cRec[36] = "検査コード023_医療機関側検査値"
		cRec[37] = "検査コード027_医療機関側検査値"
		cRec[38] = "検査コード025_医療機関側検査値"
		cRec[39] = "検査コード024_医療機関側検査値"
		cRec[40] = "検査コード028_医療機関側検査値"
		cRec[41] = "検査コード026_医療機関側検査値"
		cRec[44] = "検査コード039_医療機関側検査値"
		cRec[45] = "検査コード038_医療機関側検査値"
		cRec[46] = "検査コード037_医療機関側検査値"
		cRec[48] = "検査コード033_医療機関側検査値"
		cRec[49] = "検査コード034_医療機関側検査値"
		cRec[50] = "検査コード035_医療機関側検査値"
		cRec[51] = "検査コード080_医療機関側検査値"
		cRec[54] = "検査コード041_医療機関側検査値"
		cRec[55] = "検査コード079_医療機関側検査値"
		cRec[57] = "検査コード042_医療機関側検査値"
		cRec[64] = "検査コード031_医療機関側検査値"
		cRec[65] = "検査コード030_医療機関側判定結果"
		cRec[70] = "検査コード047_医療機関側検査値"
		cRec[72] = "検査コード021_医療機関側検査値"
		cRec[78] = "検査コード010_医療機関側検査値"
		cRec[79] = "検査コード011_医療機関側検査値"
		cRec[80] = "検査コード012_医療機関側検査値"
		cRec[81] = "検査コード013_医療機関側検査値"
		cRec[168] = "検査コード072_医療機関側検査値"
		cRec[176] = "検査コード049_医療機関側検査値"
		cRec[179] = "検査コード050_医療機関側検査値"
		cRec[182] = "検査コード051_医療機関側検査値"
		cRec[185] = "検査コード052_医療機関側検査値"
		cRec[186] = "検査コード053_医療機関側検査値"
		cRec[187] = "検査コード054_医療機関側検査値"
		cRec[188] = "検査コード055_医療機関側検査値"
		cRec[189] = "検査コード056_医療機関側検査値"
		cRec[190] = "検査コード057_医療機関側検査値"
		cRec[191] = "検査コード058_医療機関側検査値"
		cRec[192] = "検査コード059_医療機関側検査値"
		cRec[193] = "検査コード060_医療機関側検査値"
		cRec[194] = "検査コード061_医療機関側検査値"
		cRec[195] = "検査コード081_医療機関側検査値"
		cRec[196] = "検査コード062_医療機関側検査値"
		cRec[197] = "検査コード063_医療機関側検査値"
		cRec[198] = "検査コード064_医療機関側検査値"
		cRec[199] = "検査コード082_医療機関側検査値"
		cRec[200] = "検査コード065_医療機関側検査値"
		cRec[201] = "検査コード066_医療機関側検査値"
		cRec[202] = "検査コード067_医療機関側検査値"
		cRec[203] = "検査コード068_医療機関側検査値"
		cRec[204] = "検査コード069_医療機関側検査値"
		cRec[205] = "検査コード070_医療機関側検査値"
		cRec[212] = "検査コード020_医療機関側検査値"
		cRec[213] = "検査コード020_医療機関側検査値"
		cRec[214] = "検査コード044_医療機関側検査値"
		cRec[215] = "検査コード045_医療機関側検査値"
		cRec[216] = "検査コード016_医療機関側検査値"
		cRec[217] = "検査コード017_医療機関側検査値"
		cRec[218] = "検査コード018_医療機関側検査値"
		cRec[219] = "検査コード019_医療機関側検査値"
		cRec[220] = "検査コード046_医療機関側検査値"
		cRec[221] = "検査コード046_医療機関側検査値"
		//writer.Write(cRec)
		row = sheet.AddRow()
		for _, cell = range cRec {
			vcell = row.AddCell()
			vcell.Value = cell
		}
		// 3行目（タイトル）
		for I, _ = range cRec {
			cRec[I] = ""
		}

		cRec[0] = "#従業員番号"
		cRec[1] = "組合コード"
		cRec[2] = "受診者ID"
		cRec[3] = "保険証記号"
		cRec[4] = "保険証番号"
		cRec[5] = "続柄"
		cRec[6] = "枝番"
		cRec[7] = "所属コード"
		cRec[8] = "所属名称"
		cRec[9] = "社員番号"
		cRec[10] = "加入番号"
		cRec[11] = "扶養番号"
		cRec[12] = "受診者区分"
		cRec[13] = "性別"
		cRec[14] = "氏名漢字"
		cRec[15] = "氏名カナ"
		cRec[16] = "生年月日"
		cRec[17] = "実施年度"
		cRec[18] = "年齢"
		cRec[19] = "受診日"
		cRec[20] = "健診区分"
		cRec[21] = "医療機関コード"
		cRec[22] = "医療機関名称"
		cRec[23] = "機関コード"
		cRec[24] = "機関名称"
		cRec[25] = "機関住所"
		cRec[26] = "受付NO"
		cRec[27] = "身長"
		cRec[28] = "体重"
		cRec[29] = "BMI"
		cRec[30] = "内臓脂肪面積"
		cRec[31] = "腹囲"
		cRec[32] = "業務歴"
		cRec[33] = "既往歴"
		cRec[34] = "自覚症状"
		cRec[35] = "他覚症状"
		cRec[36] = "収縮期血圧(その他)"
		cRec[37] = "収縮期血圧(２回目)"
		cRec[38] = "収縮期血圧(１回目)"
		cRec[39] = "拡張期血圧(その他)"
		cRec[40] = "拡張期血圧(２回目)"
		cRec[41] = "拡張期血圧(１回目)"
		cRec[42] = "採血時間"
		cRec[43] = "総コレステロール"
		cRec[44] = "中性脂肪"
		cRec[45] = "HDLコレステロール"
		cRec[46] = "LDLコレステロール"
		cRec[47] = "NON-HDLコレステロール"
		cRec[48] = "GOT(AST)"
		cRec[49] = "GPT(ALT)"
		cRec[50] = "γ-GT(γ-GTP)"
		cRec[51] = "血清クレアチニン"
		cRec[52] = "eGFR"
		cRec[53] = "血清尿酸"
		cRec[54] = "空腹時血糖"
		cRec[55] = "随時血糖"
		cRec[56] = "HbA1c"
		cRec[57] = "HbA1c(NGSP)"
		cRec[58] = "尿糖"
		cRec[59] = "尿蛋白"
		cRec[60] = "尿潜血"
		cRec[61] = "尿素窒素"
		cRec[62] = "尿ウロビリノーゲン"
		cRec[63] = "ヘマトクリット値"
		cRec[64] = "血色素量(ヘモグロビン値)"
		cRec[65] = "赤血球数"
		cRec[66] = "貧血検査実施理由"
		cRec[67] = "白血球数"
		cRec[68] = "血小板数"
		cRec[69] = "血清アミラーゼ"
		cRec[70] = "心電図(所見)"
		cRec[71] = "心電図(実施理由)"
		cRec[72] = "胸部X線検査(所見)"
		cRec[73] = "胸部X線検査(撮影年月日)"
		cRec[74] = "喀痰検査(塗抹鏡検 一般細菌)(所見)"
		cRec[75] = "喀痰検査(塗抹鏡検 抗酸菌)"
		cRec[76] = "喀痰検査(ガフキー号数)"
		cRec[77] = "便潜血"
		cRec[78] = "視力(裸眼右)"
		cRec[79] = "視力(矯正右)"
		cRec[80] = "視力(裸眼左)"
		cRec[81] = "視力(矯正左)"
		cRec[82] = "聴力(右1000Hz)"
		cRec[83] = "聴力(右4000Hz)"
		cRec[84] = "聴力(左1000Hz)"
		cRec[85] = "聴力(左4000Hz)"
		cRec[86] = "聴力(その他の所見)"
		cRec[87] = "眼底検査(キースワグナー分類)"
		cRec[88] = "眼底検査(シェイエ分類:H)"
		cRec[89] = "眼底検査(シェイエ分類:S)"
		cRec[90] = "眼底検査(SCOTT分類)"
		cRec[91] = "眼底検査（wong-Mitchell分類）"
		cRec[92] = "眼底検査（改変Davis分類）"
		cRec[93] = "眼底検査(その他の所見)"
		cRec[94] = "眼底検査(実施理由)"
		cRec[95] = "その他の法定特殊健康診断"
		cRec[96] = "その他の法定検査"
		cRec[97] = "その他の検査"
		cRec[98] = "追加項目1"
		cRec[99] = "追加項目2"
		cRec[100] = "追加項目3"
		cRec[101] = "追加項目4"
		cRec[102] = "追加項目5"
		cRec[103] = "追加項目6"
		cRec[104] = "追加項目7"
		cRec[105] = "追加項目8"
		cRec[106] = "追加項目9"
		cRec[107] = "追加項目10"
		cRec[108] = "BMI判定"
		cRec[109] = "内臓脂肪面積判定"
		cRec[110] = "腹囲判定"
		cRec[111] = "血圧判定"
		cRec[112] = "総コレステロール判定"
		cRec[113] = "中性脂肪判定"
		cRec[114] = "HDLコレステロール判定"
		cRec[115] = "LDLコレステロール判定"
		cRec[116] = "NON-HDLコレステロール判定"
		cRec[117] = "GOT(AST)判定"
		cRec[118] = "GPT(ALT)判定"
		cRec[119] = "γ-GT(γ-GTP)判定"
		cRec[120] = "血清クレアチニン判定"
		cRec[121] = "eGFR判定"
		cRec[122] = "血清尿酸判定"
		cRec[123] = "空腹時血糖判定"
		cRec[124] = "随時血糖判定"
		cRec[125] = "HbA1c判定"
		cRec[126] = "HbA1c（NGSP)判定"
		cRec[127] = "尿糖判定"
		cRec[128] = "尿蛋白判定"
		cRec[129] = "尿潜血判定"
		cRec[130] = "尿素窒素判定"
		cRec[131] = "尿ウロビリノーゲン判定"
		cRec[132] = "ヘマトクリット値判定"
		cRec[133] = "血色素量(ヘモグロビン値)判定"
		cRec[134] = "赤血球数判定"
		cRec[135] = "白血球数判定"
		cRec[136] = "血小板数判定"
		cRec[137] = "視力(右)判定"
		cRec[138] = "視力(左)判定"
		cRec[139] = "追加項目判定1"
		cRec[140] = "追加項目判定2"
		cRec[141] = "追加項目判定3"
		cRec[142] = "追加項目判定4"
		cRec[143] = "追加項目判定5"
		cRec[144] = "追加項目判定6"
		cRec[145] = "追加項目判定7"
		cRec[146] = "追加項目判定8"
		cRec[147] = "追加項目判定9"
		cRec[148] = "追加項目判定10"
		cRec[149] = "コメント"
		cRec[150] = "総合判定"
		cRec[151] = "受診勧奨区分"
		cRec[152] = "指導状態"
		cRec[153] = "再検査区分"
		cRec[154] = "一次健診日"
		cRec[155] = "結果通知区分"
		cRec[156] = "メタボリック判定(血圧リスク)"
		cRec[157] = "メタボリック判定(血糖リスク)"
		cRec[158] = "メタボリック判定(脂質リスク)"
		cRec[159] = "メタボリック判定(リスクカウント)"
		cRec[160] = "支援レベル(血圧リスク)"
		cRec[161] = "支援レベル(血糖リスク)"
		cRec[162] = "支援レベル(脂質リスク)"
		cRec[163] = "支援レベル(喫煙リスク)"
		cRec[164] = "支援レベル(リスクカウント)"
		cRec[165] = "メタボリックシンドローム判定"
		cRec[166] = "支援レベル"
		cRec[167] = "医師の診断(判定)"
		cRec[168] = "健康診断を実施した医師の氏名"
		cRec[169] = "医師の意見"
		cRec[170] = "意見を述べた医師の氏名"
		cRec[171] = "歯科医師による健康診断"
		cRec[172] = "歯科医師による健康診断を実施した歯科医師の氏名"
		cRec[173] = "歯科医師の意見"
		cRec[174] = "意見を述べた歯科医師の氏名"
		cRec[175] = "備考"
		cRec[176] = "服薬１_血圧"
		cRec[177] = "血圧_薬剤"
		cRec[178] = "血圧_服薬理由"
		cRec[179] = "服薬２_血糖"
		cRec[180] = "血糖_薬剤"
		cRec[181] = "血糖_服薬理由"
		cRec[182] = "服薬３_脂質"
		cRec[183] = "脂質_薬剤"
		cRec[184] = "脂質_服薬理由"
		cRec[185] = "既往歴１_脳血管"
		cRec[186] = "既往歴２_心血管"
		cRec[187] = "既往歴３_腎不全人工透析"
		cRec[188] = "貧血"
		cRec[189] = "喫煙"
		cRec[190] = "２０歳からの体重変化"
		cRec[191] = "３０分以上の運動習慣"
		cRec[192] = "歩行又は身体活動"
		cRec[193] = "歩行速度"
		cRec[194] = "１年間の体重変化"
		cRec[195] = "食事についての咀嚼"
		cRec[196] = "食べ方１_早食い等"
		cRec[197] = "食べ方２_就寝前"
		cRec[198] = "食べ方３_夜食間食"
		cRec[199] = "食べ方３_三食以外の間食"
		cRec[200] = "食習慣"
		cRec[201] = "飲酒"
		cRec[202] = "飲酒量"
		cRec[203] = "睡眠"
		cRec[204] = "生活習慣の改善"
		cRec[205] = "保健指導の希望"
		cRec[206] = "報告対象区分"
		cRec[207] = "保健指導からの除外"
		cRec[208] = "取込年月日"
		cRec[209] = "胸部X線判定①"
		cRec[210] = "胸部X線判定②"
		cRec[211] = "心電図判定"
		cRec[212] = "胸部レントゲン検査"
		cRec[213] = "胸部レントゲン判定"
		cRec[214] = "尿糖"
		cRec[215] = "尿蛋白"
		cRec[216] = "聴力(右1000Hz)"
		cRec[217] = "聴力(右4000Hz)"
		cRec[218] = "聴力(左1000Hz)"
		cRec[219] = "聴力(左4000Hz)"
		cRec[220] = "心電図検査"
		cRec[221] = "心電図判定"
		//writer.Write(cRec)
		row = sheet.AddRow()
		for _, cell = range cRec {
			vcell = row.AddCell()
			vcell.Value = cell
		}

		// 4行目移行（データ）
		r = 3
		inRecsMax := len(inRecs)
		for J := 1; J < inRecsMax; J++ {
			for I, _ = range cRec {
				cRec[I] = ""
			}

			if inRecs[J][4] == coRec[0] && coseCheck(inRecs[J][13]) {
				// 0.社員番号
				if len(inRecs[J][0]) != 10 {
					log.Printf("社員番号が10桁ではありません:%v\r\n", inRecs[J][0])
				}
				cRec[0] = inRecs[J][0]

				// 1.組合コード

				// 2.受診者ID
				cRec[2] = inRecs[J][1]

				// 3.保険証記号
				cRec[3] = inRecs[J][2]

				// 4.保険証番号
				cRec[4] = inRecs[J][3]

				// 5.続柄
				// 6.枝番
				// 7.所属コード
				cRec[7] = inRecs[J][6]

				// 8.所属名称
				cRec[8] = inRecs[J][7]

				// 9.社員番号
				cRec[9] = inRecs[J][0]

				// 10.加入番号

				// 11.扶養番号

				// 12.受診者区分

				// 13.性別
				cRec[13] = inRecs[J][8]

				// 14.氏名漢字
				cRec[14] = inRecs[J][9]

				// 15.氏名カナ
				cRec[15] = string(norm.NFKC.Bytes([]byte(inRecs[J][10])))

				// 16.生年月日
				cRec[16] = WaToSeireki(inRecs[J][11])

				// 17.実施年度
				cRec[17] = nendo(inRecs[J][15])

				// 18.年齢
				cRec[18] = inRecs[J][12]

				// 19.受診日
				cRec[19] = strings.Replace(inRecs[J][15], "-", "/", -1)

				// 20.健診区分
				cRec[20] = "事業者健診"

				// 21.医療機関コード
				cRec[21] = "013-61"

				// 22.医療機関名称
				cRec[22] = "医療法人社団　松英会"

				// 23.機関コード
				cRec[23] = "1311131242"

				// 24.機関名称
				cRec[24] = "医療法人社団　松英会"

				// 25.機関住所
				cRec[25] = "143-0027 大田区中馬込1-5-8"

				// 26.受付NO
				cRec[26] = inRecs[J][16]

				// 27.身長
				cRec[27] = inRecs[J][17]

				// 28.体重
				cRec[28] = inRecs[J][18]

				// 29.BMI
				cRec[29] = inRecs[J][19]

				// 30.内臓脂肪面積

				// 31.腹囲
				cRec[31] = inRecs[J][20]

				// 32.業務歴

				// 33.既往歴
				kiou := ""
				for k := 0; k < 10; k++ {
					kp := 21 + (k * 2)
					kiouB := kiouSet(inRecs[J][kp])
					kiouT := kiouSet(inRecs[J][kp+1])
					if kiouB != "" {
						if utf8.RuneCountInString(kiou+" "+kiouB+kiouT) > 25 {
							if utf8.RuneCountInString(kiou+" "+kiouB) > 25 {
								break
							} else {
								if kiou == "" {
									kiou = kiouB
								} else {
									kiou = kiou + " " + kiouB
								}
							}
						} else {
							if kiou == "" {
								kiou = kiouB + kiouT
							} else {
								kiou = kiou + " " + kiouB + kiouT
							}
						}
					}
				}

				cRec[33] = kiou

				// 34.自覚症状
				cRec[34] = syoken(inRecs[J][41] + " " + inRecs[J][42] + " " + inRecs[J][43])

				// 35.他覚症状
				cRec[35] = syoken(inRecs[J][44] + " " + inRecs[J][45] + " " + inRecs[J][46])

				// 36.収縮期血圧(その他)

				// 37.収縮期血圧(２回目)
				cRec[37] = inRecs[J][47]

				// 38.収縮期血圧(１回目)
				cRec[38] = inRecs[J][48]

				// 39.拡張期血圧(その他)

				// 40.拡張期血圧(２回目)
				cRec[40] = inRecs[J][49]

				// 41.拡張期血圧(１回目)
				cRec[41] = inRecs[J][50]

				// 42.採血時間

				// 43.総コレステロール
				cRec[43] = inRecs[J][53]

				// 44.中性脂肪
				cRec[44] = inRecs[J][54]

				// 45.HDLコレステロール
				cRec[45] = inRecs[J][55]

				// 46.LDLコレステロール
				cRec[46] = inRecs[J][56]

				// 47.NON-HDLコレステロール
				cRec[47] = inRecs[J][57]

				// 48.GOT(AST)
				cRec[48] = inRecs[J][58]

				// 49.GPT(ALT)
				cRec[49] = inRecs[J][59]

				// 50.γ-GT(γ-GTP)
				cRec[50] = inRecs[J][60]

				// 51.血清クレアチニン
				cRec[51] = inRecs[J][61]

				// 52.eGFR
				cRec[52] = inRecs[J][62]

				// 53.血清尿酸
				cRec[53] = inRecs[J][63]

				// 54.空腹時血糖
				// 55.随時血糖
				if syokugo(inRecs[J][51], inRecs[J][52]) {
					cRec[55] = inRecs[J][64]
				} else {
					cRec[54] = inRecs[J][64]
				}

				// 56.HbA1c
				// 57.HbA1c(NGSP)
				cRec[57] = inRecs[J][65]

				// 58.尿糖
				cRec[58] = nyouT(inRecs[J][66])

				// 59.尿蛋白
				cRec[59] = nyouT(inRecs[J][67])

				// 60.尿潜血
				cRec[60] = inRecs[J][68]

				// 61.尿素窒素
				cRec[61] = inRecs[J][69]

				// 62.尿ウロビリノーゲン
				cRec[62] = inRecs[J][70]

				// 63.ヘマトクリット値
				cRec[63] = inRecs[J][71]

				// 64.血色素量(ヘモグロビン値)
				cRec[64] = inRecs[J][72]

				// 65.赤血球数
				cRec[65] = inRecs[J][73]

				// 66.貧血検査実施理由

				// 67.白血球数
				cRec[67] = inRecs[J][74]

				// 68.血小板数
				cRec[68] = inRecs[J][75]

				// 69.血清アミラーゼ
				cRec[69] = inRecs[J][76]

				// 70.心電図(所見)
				cRec[70] = syoken(inRecs[J][78] + " " + inRecs[J][79] + " " + inRecs[J][80] + " " + inRecs[J][81])

				// 71.心電図(実施理由)

				// 72.胸部X線検査(所見)
				cRec[72] = syoken(inRecs[J][83] + " " + inRecs[J][84] + " " + inRecs[J][85])

				// 73.胸部X線検査(撮影年月日)
				if inRecs[J][82] != "" {
					cRec[73] = strings.Replace(inRecs[J][15], "-", "/", -1)
				}

				// 74.喀痰検査(塗抹鏡検 一般細菌)(所見)

				// 75.喀痰検査(塗抹鏡検 抗酸菌)

				// 76.喀痰検査(ガフキー号数)

				// 77.便潜血
				cRec[77] = inRecs[J][86]

				// 78.視力(裸眼右)
				cRec[78] = eye(inRecs[J][87])

				// 79.視力(矯正右)
				cRec[79] = eye(inRecs[J][88])

				// 80.視力(裸眼左)
				cRec[80] = eye(inRecs[J][89])

				// 81.視力(矯正左)
				cRec[81] = eye(inRecs[J][90])

				// 82.聴力(右1000Hz)
				cRec[82] = syokenumu(inRecs[J][99])

				// 83.聴力(右4000Hz)
				cRec[83] = syokenumu4k(inRecs[J][101], inRecs[J][103])

				// 84.聴力(左1000Hz)
				cRec[84] = syokenumu(inRecs[J][100])

				// 85.聴力(左4000Hz)
				cRec[85] = syokenumu4k(inRecs[J][102], inRecs[J][104])

				// 86.聴力(その他の所見)

				// 87.眼底検査(キースワグナー分類)

				// 88.眼底検査(シェイエ分類:H)

				// 89.眼底検査(シェイエ分類:S)

				// 90.眼底検査(SCOTT分類)

				// 91.眼底検査(wong-Mitchell分類)

				// 92.眼底検査(改変Davis分類)

				// 93.眼底検査(その他の所見)

				// 94.眼底検査(実施理由)

				// 95.その他の法定特殊健康診断

				// 96.その他の法定検査

				// 97.その他の検査

				// 98.追加項目1

				// 99.追加項目2

				// 100.追加項目3

				// 101.追加項目4

				// 102.追加項目5

				// 103.追加項目6

				// 104.追加項目7

				// 105.追加項目8

				// 106.追加項目9

				// 107.追加項目10

				// 108.BMI判定
				cRec[108] = inRecs[J][105]

				// 109.内臓脂肪面積判定

				// 110.腹囲判定
				cRec[110] = inRecs[J][106]

				// 111.血圧判定
				cRec[111] = string(norm.NFKC.Bytes([]byte(inRecs[J][107])))

				// 112.総コレステロール判定
				cRec[112] = inRecs[J][108]

				// 113.中性脂肪判定
				cRec[113] = inRecs[J][109]

				// 114.HDLコレステロール判定
				cRec[114] = inRecs[J][110]

				// 115.LDLコレステロール判定
				cRec[115] = inRecs[J][111]

				// 116.NON-HDLコレステロール判定
				cRec[116] = inRecs[J][112]

				// 117.GOT(AST)判定
				cRec[117] = inRecs[J][113]

				// 118.GPT(ALT)判定
				cRec[118] = inRecs[J][114]

				// 119.γ-GT(γ-GTP)判定
				cRec[119] = inRecs[J][115]

				// 120.血清クレアチニン判定
				cRec[120] = inRecs[J][116]

				// 121.eGFR判定
				cRec[121] = inRecs[J][117]

				// 122.血清尿酸判定
				cRec[122] = inRecs[J][118]

				// 123.空腹時血糖判定
				// 124.随時血糖判定
				if syokugo(inRecs[J][51], inRecs[J][52]) {
					cRec[124] = toH(inRecs[J][64])
				} else {
					cRec[123] = inRecs[J][119]
				}

				// 125.HbA1c判定

				// 126.HbA1c（NGSP)判定
				cRec[126] = inRecs[J][120]

				// 127.尿糖判定
				cRec[127] = inRecs[J][121]

				// 128.尿蛋白判定
				cRec[128] = inRecs[J][122]

				// 129.尿潜血判定
				cRec[129] = inRecs[J][123]

				// 130.尿素窒素判定
				cRec[130] = inRecs[J][124]

				// 131.尿ウロビリノーゲン判定
				cRec[131] = inRecs[J][125]

				// 132.ヘマトクリット値判定
				cRec[132] = inRecs[J][126]

				// 133.血色素量(ヘモグロビン値)判定
				cRec[133] = inRecs[J][127]

				// 134.赤血球数判定
				cRec[134] = inRecs[J][128]

				// 135.白血球数判定
				cRec[135] = inRecs[J][129]

				// 136.血小板数判定
				cRec[136] = inRecs[J][130]

				// 137.視力(右)判定
				cRec[137] = eyeHantei(inRecs[J][131], inRecs[J][132])

				// 138.視力(左)判定
				cRec[138] = eyeHantei(inRecs[J][133], inRecs[J][134])

				// 139.追加項目判定1

				// 140.追加項目判定2

				// 141.追加項目判定3

				// 142.追加項目判定4

				// 143.追加項目判定5

				// 144.追加項目判定6

				// 145.追加項目判定7

				// 146.追加項目判定8

				// 147.追加項目判定9

				// 148.追加項目判定10

				// 149.コメント

				// 150.総合判定
				if inRecs[J][135] == "" {
					log.Print("総合判定が抜けている方がいます。")
				}
				cRec[150] = inRecs[J][135]

				// 151.受診勧奨区分

				// 152.指導状態

				// 153.再検査区分

				// 154.一次健診日

				// 155.結果通知区分

				// 156.メタボリック判定(血圧リスク)

				// 157.メタボリック判定(血糖リスク)

				// 158.メタボリック判定(脂質リスク)

				// 159.メタボリック判定(リスクカウント)

				// 160.支援レベル(血圧リスク)

				// 161.支援レベル(血糖リスク)

				// 162.支援レベル(脂質リスク)

				// 163.支援レベル(喫煙リスク)

				// 164.支援レベル(リスクカウント)

				// 165.メタボリックシンドローム判定
				cRec[165] = inRecs[J][137]

				// 166.支援レベル
				cRec[166] = inRecs[J][138]

				// 167.医師の診断(判定)
				cRec[167] = inRecs[J][136]

				// 168.健康診断を実施した医師の氏名
				cRec[168] = "寺門　節雄"

				// 169.医師の意見

				// 170.意見を述べた医師の氏名

				// 171.歯科医師による健康診断

				// 172.歯科医師による健康診断を実施した歯科医師の氏名

				// 173.歯科医師の意見

				// 174.意見を述べた歯科医師の氏名

				// 175.備考

				// 176.服薬１_血圧
				cRec[176] = inRecs[J][139]

				// 177.血圧_薬剤

				// 178.血圧_服薬理由

				// 179.服薬２_血糖
				cRec[179] = inRecs[J][140]

				// 180.血糖_薬剤

				// 181.血糖_服薬理由

				// 182.服薬３_脂質
				cRec[182] = inRecs[J][141]

				// 183.脂質_薬剤

				// 184.脂質_服薬理由

				// 185.既往歴１_脳血管
				cRec[185] = inRecs[J][142]

				// 186.既往歴２_心血管
				cRec[186] = inRecs[J][143]

				// 187.既往歴３_腎不全人工透析
				cRec[187] = inRecs[J][144]

				// 188.貧血
				cRec[188] = inRecs[J][145]

				// 189.喫煙
				cRec[189] = inRecs[J][146]

				// 190.２０歳からの体重変化
				cRec[190] = inRecs[J][147]

				// 191.３０分以上の運動習慣
				cRec[191] = inRecs[J][148]

				// 192.歩行又は身体活動
				cRec[192] = inRecs[J][149]

				// 193.歩行速度
				cRec[193] = inRecs[J][150]

				// 194.１年間の体重変化
				//cRec[194] = inRecs[J][151]

				// 195.食事についての咀嚼
				cRec[195] = inRecs[J][151]

				// 196.食べ方１_早食い等
				cRec[196] = inRecs[J][152]

				// 197.食べ方２_就寝前
				cRec[197] = inRecs[J][153]

				// 198.食べ方３_夜食間食
				//cRec[198] = inRecs[J][154]

				// 199.食べ方３_三食以外の間食
				cRec[199] = inRecs[J][154]

				// 200.食習慣
				cRec[200] = inRecs[J][155]

				// 201.飲酒
				cRec[201] = inRecs[J][156]

				// 202.飲酒量
				cRec[202] = inRecs[J][157]

				// 203.睡眠
				cRec[203] = inRecs[J][158]

				// 204.生活習慣の改善
				cRec[204] = inRecs[J][159]

				// 205.保健指導の希望
				cRec[205] = inRecs[J][160]

				// 206.報告対象区分

				// 207.保健指導からの除外

				// 208.取込年月日

				// 209.胸部X線判定①
				cRec[209] = string(norm.NFKC.Bytes([]byte(inRecs[J][82])))

				// 210.胸部X線判定②
				cRec[210] = string(norm.NFKC.Bytes([]byte(inRecs[J][82])))

				// 211.心電図判定
				cRec[211] = string(norm.NFKC.Bytes([]byte(inRecs[J][77])))

				// 212.胸部レントゲン検査
				cRec[212] = hanteiCode(string(norm.NFKC.Bytes([]byte(inRecs[J][82]))))

				// 213.胸部レントゲン判定
				cRec[213] = hanteiCode(string(norm.NFKC.Bytes([]byte(inRecs[J][82]))))

				// 214.尿糖
				cRec[214] = nyou(inRecs[J][66])

				// 215.尿蛋白
				cRec[215] = nyou(inRecs[J][67])

				// 216.聴力(右1000Hz)
				cRec[216] = syokenumuCode(syokenumu(inRecs[J][99]))

				// 217.聴力(右4000Hz)
				cRec[217] = syokenumuCode(syokenumu4k(inRecs[J][101], inRecs[J][103]))

				// 218.聴力(左1000Hz)
				cRec[218] = syokenumuCode(syokenumu(inRecs[J][100]))

				// 219.聴力(左4000Hz)
				cRec[219] = syokenumuCode(syokenumu4k(inRecs[J][102], inRecs[J][104]))

				// 220.心電図検査
				cRec[220] = hanteiCode(string(norm.NFKC.Bytes([]byte(inRecs[J][77]))))

				// 221.心電図判定
				cRec[221] = hanteiCode(string(norm.NFKC.Bytes([]byte(inRecs[J][77]))))

				//writer.Write(cRec)
				row = sheet.AddRow()
				for _, cell = range cRec {
					// sheet.Cell(r, c).Value = cell
					vcell = row.AddCell()
					vcell.Value = cell
				}
				r++
			}
		}

		//writer.Flush()
		err = excelFile.Save(excelName)
		failOnError(err)
	}

}

func meiboCreate(filename string, inRecs [][]string, coRecs [][]string) {
	var excelFile *xlsx.File
	var sheet *xlsx.Sheet
	var vcell *xlsx.Cell
	var err error

	recLen := 14 //出力するレコードの項目数
	jRec := make([]string, recLen)

	day := time.Now()

	//会社毎に受診者名簿を作成する
	for _, coRec := range coRecs {

		/*
			outfile, err := os.Create(filename + coRec[1] + "受診者名簿" + day.Format("20060102") + ".txt")
			failOnError(err)
			defer outfile.Close()

			writer := csv.NewWriter(transform.NewWriter(outfile, japanese.ShiftJIS.NewEncoder()))
			writer.Comma = '\t'
			writer.UseCRLF = true
		*/

		excelName := filename + coRec[1] + "受診者名簿" + day.Format("20060102") + ".xlsx"
		excelFile = xlsx.NewFile()
		xlsx.SetDefaultFont(11, "ＭＳ Ｐゴシック")
		sheet, err = excelFile.AddSheet("データ")
		failOnError(err)

		r := 0
		for _, inRec := range inRecs {
			if (coRec[0] == inRec[4] && coseCheck(inRec[13])) || inRec[4] == "所属cd１" {
				jRec[0] = inRec[4]
				jRec[1] = inRec[5]
				jRec[2] = inRec[6]
				jRec[3] = inRec[7]
				jRec[4] = inRec[13]
				jRec[5] = inRec[14]
				jRec[6] = inRec[0]
				jRec[7] = inRec[10]
				jRec[8] = inRec[9]
				jRec[9] = inRec[8]
				jRec[10] = inRec[11]
				jRec[11] = inRec[12]
				jRec[12] = inRec[15]
				jRec[13] = inRec[16]
				//writer.Write(jRec)
				row := sheet.AddRow()
				for _, cell := range jRec {
					// sheet.Cell(r, c).Value = cell
					vcell = row.AddCell()
					vcell.Value = cell

				}
				r++
			}

		}

		//writer.Flush()
		err = excelFile.Save(excelName)
		failOnError(err)
	}

}

func WaToSeireki(nen string) string {

	if len(nen) != 9 {
		return nen
	} else {
		w := nen[0:1]
		y := nen[1 : 1+2]
		yi, _ := strconv.Atoi(y)
		m := nen[4 : 4+2]
		d := nen[7 : 7+2]

		switch w {
		case "M":
			yi = yi + 1867
		case "T":
			yi = yi + 1911
		case "S":
			yi = yi + 1925
		case "H":
			yi = yi + 1988
		default:
			yi = 0
		}

		if yi == 0 {
			return "err"
		} else {
			return fmt.Sprint(yi) + "/" + m + "/" + d
		}
	}
}

func nendo(JDay string) string {
	var nen int
	t, _ := time.Parse("2006-01-02", JDay)
	if t.Month() > 3 {
		nen = t.Year()
	} else {
		nen = t.Year() - 1
	}

	return strconv.Itoa(nen)
}

func kiouSet(s string) string {
	var spos, epos int
	//全角記号を半角へ
	s = strings.Replace(s, "（", "(", -1)
	s = strings.Replace(s, "）", ")", -1)
	s = strings.Replace(s, "　", " ", -1)

	// ()でくくった文字は削除
	for {
		spos = strings.LastIndex(s, "(")
		epos = strings.LastIndex(s, ")")

		if epos == -1 {
			break
		} else if spos == -1 {
			break
		} else {
			//log.Print(s + ":epos→" + fmt.Sprint(epos) + " len→" + fmt.Sprint(len(s)) + "\r\n")
			s = s[:spos] + s[epos+1:]
		}
	}

	// 余分なスペースを削除
	s = dsTrim(s)
	s = strings.Trim(s, " ")

	return s
}

func dsTrim(s string) string {
	for {
		if strings.Contains(s, "  ") {
			s = strings.Replace(s, "  ", " ", -1)
		} else {
			return s
		}
	}
}

func cutStrings(s string, maxLen int) string {
	s = string([]rune(s)[:maxLen])
	return s
}

func syoken(s string) string {
	s = strings.Replace(s, "　", " ", -1)
	s = strings.Trim(s, " ")

	for {
		if utf8.RuneCountInString(s) > 25 {
			pos := strings.LastIndex(s, " ")
			s = s[:pos]
		} else {
			break
		}
	}

	return s
}

func nyou(s string) string {

	switch s {
	case "":
		s = ""
	case "－":
		s = "1"
	case "+-":
		s = "2"
	case "＋":
		s = "3"
	case "2+":
		s = "4"
	case "3+":
		s = "5"
	case "4+":
		s = "6"
	case "5+":
		s = "6"
	default:
		s = "err"
	}
	return s
}

func nyouT(s string) string {

	switch s {
	case "":
		s = ""
	case "－":
		s = "－"
	case "+-":
		s = "±"
	case "＋":
		s = "+"
	case "2+":
		s = "++"
	case "3+":
		s = "+++"
	case "4+":
		s = "++++"
	case "5+":
		s = "++++"
	default:
		s = "err"
	}
	return s
}

func eye(s string) string {
	if s == "0.1↓" {
		s = "0.0"
	}
	return s
}

func syokenumu(s string) string {

	switch s {
	case "":
		s = ""
	case "A":
		s = "所見なし"
	case "B":
		s = "所見あり"
	case "C":
		s = "所見あり"
	case "D":
		s = "所見あり"
	case "E":
		s = "所見あり"
	case "F":
		s = "所見あり"
	case "G":
		s = "所見あり"
	default:
		s = "err"
	}
	return s
}

func syokenumu4k(s1, s2 string) string {
	var s string
	s1s := syokenumu(s1)
	s2s := syokenumu(s2)

	s = s1s
	if s == "" {
		s = s2s
	}

	return s
}

func syokenumuCode(s string) string {

	if s == "所見なし" {
		return "1"
	} else if s == "所見あり" {
		return "2"
	} else {
		return s
	}

}

func eyeHantei(s1, s2 string) string {
	var s string

	switch s1 {
	case "":
		s = s2
	case "A":
		s = "A"
	case "B":
		if s2 == "A" {
			s = s2
		} else {
			s = "B"
		}
	case "C":
		if s2 == "A" || s2 == "B" {
			s = s2
		} else {
			s = "C"
		}
	case "D":
		if s2 == "A" || s2 == "B" || s2 == "C" {
			s = s2
		} else {
			s = "D"
		}
	case "E":
		if s2 == "A" || s2 == "B" || s2 == "C" || s2 == "D" {
			s = s2
		} else {
			s = "E"
		}
	case "F":
		if s2 == "A" || s2 == "B" || s2 == "C" || s2 == "D" || s2 == "E" {
			s = s2
		} else {
			s = "F"
		}
	case "G":
		if s2 == "A" || s2 == "B" || s2 == "C" || s2 == "D" || s2 == "E" || s2 == "F" {
			s = s2
		} else {
			s = "G"
		}
	default:
		s = "err"
	}

	return s
}

func hanteiCode(s string) string {

	switch s {
	case "":
		s = ""
	case "A":
		s = "1"
	case "B":
		s = "2"
	case "C":
		s = "3"
	case "D":
		s = "4"
	case "E":
		s = "5"
	case "F":
		s = "6"
	case "G":
		s = "7"
	default:
		s = "err"
	}
	return s
}

func toH(s string) string {

	v := ""
	i, _ := strconv.Atoi(s)
	if s == "" {
		v = ""
	} else if i <= 59 {
		v = "E"
	} else if (i >= 60) && (i <= 69) {
		v = "C"
	} else if (i >= 70) && (i <= 109) {
		v = "A"
	} else if (i >= 110) && (i <= 139) {
		v = "B"
	} else if (i >= 140) && (i <= 199) {
		v = "E"
	} else if i >= 200 {
		v = "F"
	}

	return v

}

func syokugo(t, h string) bool {
	hh, _ := strconv.ParseFloat(h, 32)

	if (t == "とった") && (hh <= 2.0) {
		return true
	} else {
		return false
	}

}

func coseCheck(cose string) bool {
	// 定健コースかチェックする。20001001000001_ﾄﾖﾀ_34才以下,20001001000002_ﾄﾖﾀ_35才以上,20001001000003_ﾄﾖﾀ_関連35才以上,20001001000007_ﾄﾖﾀ_関連35才以上_便潜血
	// 雇い入れ時健診追加。20001001000005トヨタ_雇入時
	// 人間ドックデータ追加 95001001000401 ﾄﾖﾀ販売_すこやか,95001001000402 ﾄﾖﾀ販売_人間ドック
	coses := []string{"20001001000001", "20001001000002", "20001001000003", "20001001000007", "20001001000005", "95001001000401", "95001001000402"}

	for _, chkcose := range coses {
		if cose == chkcose {
			return true
		}
	}

	return false

}
