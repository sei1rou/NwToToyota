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
	companys := [][]string{{"2000100100000001", "東京トヨペット株式会社", "0"},
		{"2000100100000002", "ネッツトヨタ東京株式会社", "0"},
		{"2000100100000003", "ＤＵＯ東京株式会社", "0"},
		{"2000100100000004", "トヨタアドミニスタ株式会社", "0"},
		{"2000100100000005", "東京トヨタ自動車株式会社", "0"},
		{"2000100100000911", "株式会社　トヨテック", "0"},
		{"2000100100000991", "株式会社　センチュリーサービス", "0"},
		{"2000100100000992", "東京トヨタカーライフサービス㈱", "0"},
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
	outDirPlus := outDir + "/トヨタ東京販売ホールディングス" + day.Format("20060102")

	if err := os.Mkdir(outDirPlus, 0777); err != nil {
		log.Print(outDirPlus + "\r\n")
		log.Print("出力先のディレクトリを作成できませんでした\r\n")
		return outDir
	} else {
		return outDirPlus + "/"
	}
}

func dataConversion(filename string, inRecs [][]string, coRecs [][]string) {
	var excelFile *xlsx.File
	var sheet *xlsx.Sheet
	var err error
	var r int
	var c int
	var cell string

	recLen := 213 //出力するレコードの項目数
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
		excelFile = xlsx.NewFile()
		xlsx.SetDefaultFont(11, "ＭＳ Ｐゴシック")
		sheet, err = excelFile.AddSheet("データ")
		failOnError(err)

		// 1行目（タイトル）
		for I, _ = range cRec {
			cRec[I] = ""
		}

		cRec[0] = "idou.sya_bg"
		cRec[18] = "knk_kenkork.jushin_date"
		cRec[21] = "knk_kenkork_kensa.kensa_val_071"
		cRec[26] = "knk_kenkork_kensa.kensa_val_005"
		cRec[27] = "knk_kenkork_kensa.kensa_val_006"
		cRec[28] = "knk_kenkork_kensa.kensa_val_007"
		cRec[30] = "knk_kenkork_kensa.kensa_val_008"
		cRec[32] = "knk_kenkork_kensa.kensa_val_001"
		cRec[33] = "knk_kenkork_kensa.kensa_val_002"
		cRec[34] = "knk_kenkork_kensa.kensa_val_003"
		cRec[35] = "knk_kenkork_kensa.kensa_val_023"
		cRec[36] = "knk_kenkork_kensa.kensa_val_027"
		cRec[37] = "knk_kenkork_kensa.kensa_val_025"
		cRec[38] = "knk_kenkork_kensa.kensa_val_024"
		cRec[39] = "knk_kenkork_kensa.kensa_val_028"
		cRec[40] = "knk_kenkork_kensa.kensa_val_026"
		cRec[43] = "knk_kenkork_kensa.kensa_val_039"
		cRec[44] = "knk_kenkork_kensa.kensa_val_038"
		cRec[45] = "knk_kenkork_kensa.kensa_val_037"
		cRec[46] = "knk_kenkork_kensa.kensa_val_033"
		cRec[47] = "knk_kenkork_kensa.kensa_val_034"
		cRec[48] = "knk_kenkork_kensa.kensa_val_035"
		cRec[51] = "knk_kenkork_kensa.kensa_val_041"
		cRec[52] = "knk_kenkork_kensa.kensa_val_079"
		cRec[54] = "knk_kenkork_kensa.kensa_val_042"
		cRec[61] = "knk_kenkork_kensa.kensa_val_031"
		cRec[62] = "knk_kenkork_kensa.kensa_val_030"
		cRec[67] = "knk_kenkork_kensa.kensa_val_047"
		cRec[69] = "knk_kenkork_kensa.kensa_val_021"
		cRec[75] = "knk_kenkork_kensa.kensa_val_010"
		cRec[76] = "knk_kenkork_kensa.kensa_val_011"
		cRec[77] = "knk_kenkork_kensa.kensa_val_012"
		cRec[78] = "knk_kenkork_kensa.kensa_val_013"
		cRec[161] = "knk_kenkork_kensa.kensa_val_072"
		cRec[169] = "knk_kenkork_kensa.kensa_val_049"
		cRec[172] = "knk_kenkork_kensa.kensa_val_050"
		cRec[175] = "knk_kenkork_kensa.kensa_val_051"
		cRec[178] = "knk_kenkork_kensa.kensa_val_052"
		cRec[179] = "knk_kenkork_kensa.kensa_val_053"
		cRec[180] = "knk_kenkork_kensa.kensa_val_054"
		cRec[181] = "knk_kenkork_kensa.kensa_val_055"
		cRec[182] = "knk_kenkork_kensa.kensa_val_056"
		cRec[183] = "knk_kenkork_kensa.kensa_val_057"
		cRec[184] = "knk_kenkork_kensa.kensa_val_058"
		cRec[185] = "knk_kenkork_kensa.kensa_val_059"
		cRec[186] = "knk_kenkork_kensa.kensa_val_060"
		cRec[187] = "knk_kenkork_kensa.kensa_val_061"
		cRec[188] = "knk_kenkork_kensa.kensa_val_062"
		cRec[189] = "knk_kenkork_kensa.kensa_val_063"
		cRec[190] = "knk_kenkork_kensa.kensa_val_064"
		cRec[191] = "knk_kenkork_kensa.kensa_val_065"
		cRec[192] = "knk_kenkork_kensa.kensa_val_066"
		cRec[193] = "knk_kenkork_kensa.kensa_val_067"
		cRec[194] = "knk_kenkork_kensa.kensa_val_068"
		cRec[195] = "knk_kenkork_kensa.kensa_val_069"
		cRec[196] = "knk_kenkork_kensa.kensa_val_070"
		cRec[203] = "knk_kenkork_kensa.kensa_val_020"
		cRec[204] = "knk_kenkork_kensa.hantei_val_020"
		cRec[205] = "knk_kenkork_kensa.kensa_val_044"
		cRec[206] = "knk_kenkork_kensa.kensa_val_045"
		cRec[207] = "knk_kenkork_kensa.kensa_val_016"
		cRec[208] = "knk_kenkork_kensa.kensa_val_017"
		cRec[209] = "knk_kenkork_kensa.kensa_val_018"
		cRec[210] = "knk_kenkork_kensa.kensa_val_019"
		cRec[211] = "knk_kenkork_kensa.kensa_val_046"
		cRec[212] = "knk_kenkork_kensa.hantei_val_046"
		//writer.Write(cRec)
		for c, cell = range cRec {
			sheet.Cell(0, c).Value = cell
		}

		// 2行目（タイトル）
		for I, _ = range cRec {
			cRec[I] = ""
		}

		cRec[0] = "社員番号"
		cRec[18] = "受診日付"
		cRec[21] = "検査コード071_医療機関側判定結果"
		cRec[26] = "検査コード005_医療機関側検査値"
		cRec[27] = "検査コード006_医療機関側検査値"
		cRec[28] = "検査コード007_医療機関側検査値"
		cRec[30] = "検査コード008_医療機関側検査値"
		cRec[32] = "検査コード001_医療機関側検査値"
		cRec[33] = "検査コード02_医療機関側判定結果"
		cRec[34] = "検査コード003_医療機関側検査値"
		cRec[35] = "検査コード023_医療機関側検査値"
		cRec[36] = "検査コード027_医療機関側検査値"
		cRec[37] = "検査コード025_医療機関側検査値"
		cRec[38] = "検査コード024_医療機関側検査値"
		cRec[39] = "検査コード028_医療機関側検査値"
		cRec[40] = "検査コード026_医療機関側検査値"
		cRec[43] = "検査コード039_医療機関側検査値"
		cRec[44] = "検査コード038_医療機関側検査値"
		cRec[45] = "検査コード037_医療機関側検査値"
		cRec[46] = "検査コード033_医療機関側検査値"
		cRec[47] = "検査コード034_医療機関側検査値"
		cRec[48] = "検査コード035_医療機関側検査値"
		cRec[51] = "検査コード041_医療機関側検査値"
		cRec[52] = "検査コード079_医療機関側検査値"
		cRec[54] = "検査コード042_医療機関側検査値"
		cRec[61] = "検査コード031_医療機関側検査値"
		cRec[62] = "検査コード030_医療機関側判定結果"
		cRec[67] = "検査コード047_医療機関側検査値"
		cRec[69] = "検査コード021_医療機関側検査値"
		cRec[75] = "検査コード010_医療機関側検査値"
		cRec[76] = "検査コード011_医療機関側検査値"
		cRec[77] = "検査コード012_医療機関側検査値"
		cRec[78] = "検査コード013_医療機関側検査値"
		cRec[161] = "検査コード072_医療機関側検査値"
		cRec[169] = "検査コード049_医療機関側検査値"
		cRec[172] = "検査コード050_医療機関側検査値"
		cRec[175] = "検査コード051_医療機関側検査値"
		cRec[178] = "検査コード052_医療機関側検査値"
		cRec[179] = "検査コード053_医療機関側検査値"
		cRec[180] = "検査コード054_医療機関側検査値"
		cRec[181] = "検査コード055_医療機関側検査値"
		cRec[182] = "検査コード056_医療機関側検査値"
		cRec[183] = "検査コード057_医療機関側検査値"
		cRec[184] = "検査コード058_医療機関側検査値"
		cRec[185] = "検査コード059_医療機関側検査値"
		cRec[186] = "検査コード060_医療機関側検査値"
		cRec[187] = "検査コード061_医療機関側検査値"
		cRec[188] = "検査コード062_医療機関側検査値"
		cRec[189] = "検査コード063_医療機関側検査値"
		cRec[190] = "検査コード064_医療機関側検査値"
		cRec[191] = "検査コード065_医療機関側検査値"
		cRec[192] = "検査コード066_医療機関側検査値"
		cRec[193] = "検査コード067_医療機関側検査値"
		cRec[194] = "検査コード068_医療機関側検査値"
		cRec[195] = "検査コード069_医療機関側検査値"
		cRec[196] = "検査コード070_医療機関側検査値"
		cRec[203] = "検査コード020_医療機関側検査値"
		cRec[204] = "検査コード020_医療機関側検査値"
		cRec[205] = "検査コード044_医療機関側検査値"
		cRec[206] = "検査コード045_医療機関側検査値"
		cRec[207] = "検査コード016_医療機関側検査値"
		cRec[208] = "検査コード017_医療機関側検査値"
		cRec[209] = "検査コード018_医療機関側検査値"
		cRec[210] = "検査コード019_医療機関側検査値"
		cRec[211] = "検査コード046_医療機関側検査値"
		cRec[212] = "検査コード046_医療機関側検査値"
		//writer.Write(cRec)
		for c, cell = range cRec {
			sheet.Cell(1, c).Value = cell
		}
		// 3行目（タイトル）
		for I, _ = range cRec {
			cRec[I] = ""
		}

		cRec[0] = "社員番号"
		cRec[1] = "組合コード"
		cRec[2] = "受診者ID"
		cRec[3] = "保険証記号"
		cRec[4] = "保険証番号"
		cRec[5] = "続柄"
		cRec[6] = "枝番"
		cRec[7] = "所属コード"
		cRec[8] = "所属名称"
		cRec[9] = "加入番号"
		cRec[10] = "扶養番号"
		cRec[11] = "受診者区分"
		cRec[12] = "性別"
		cRec[13] = "氏名漢字"
		cRec[14] = "氏名カナ"
		cRec[15] = "生年月日"
		cRec[16] = "実施年度"
		cRec[17] = "年齢"
		cRec[18] = "受診日"
		cRec[19] = "健診区分"
		cRec[20] = "医療機関コード"
		cRec[21] = "医療機関名称"
		cRec[22] = "機関コード"
		cRec[23] = "機関名称"
		cRec[24] = "機関住所"
		cRec[25] = "受付NO"
		cRec[26] = "身長"
		cRec[27] = "体重"
		cRec[28] = "BMI"
		cRec[29] = "内臓脂肪面積"
		cRec[30] = "腹囲"
		cRec[31] = "業務歴"
		cRec[32] = "既往歴"
		cRec[33] = "自覚症状"
		cRec[34] = "他覚症状"
		cRec[35] = "収縮期血圧(その他)"
		cRec[36] = "収縮期血圧(２回目)"
		cRec[37] = "収縮期血圧(１回目)"
		cRec[38] = "拡張期血圧(その他)"
		cRec[39] = "拡張期血圧(２回目)"
		cRec[40] = "拡張期血圧(１回目)"
		cRec[41] = "採血時間"
		cRec[42] = "総コレステロール"
		cRec[43] = "中性脂肪"
		cRec[44] = "HDLコレステロール"
		cRec[45] = "LDLコレステロール"
		cRec[46] = "GOT(AST)"
		cRec[47] = "GPT(ALT)"
		cRec[48] = "γ-GT(γ-GTP)"
		cRec[49] = "血清クレアチニン"
		cRec[50] = "血清尿酸"
		cRec[51] = "空腹時血糖"
		cRec[52] = "随時血糖"
		cRec[53] = "HbA1c"
		cRec[54] = "HbA1c(NGSP)"
		cRec[55] = "尿糖"
		cRec[56] = "尿蛋白"
		cRec[57] = "尿潜血"
		cRec[58] = "尿素窒素"
		cRec[59] = "尿ウロビリノーゲン"
		cRec[60] = "ヘマトクリット値"
		cRec[61] = "血色素量(ヘモグロビン値)"
		cRec[62] = "赤血球数"
		cRec[63] = "貧血検査実施理由"
		cRec[64] = "白血球数"
		cRec[65] = "血小板数"
		cRec[66] = "血清アミラーゼ"
		cRec[67] = "心電図(所見)"
		cRec[68] = "心電図(実施理由)"
		cRec[69] = "胸部X線検査(所見)"
		cRec[70] = "胸部X線検査(撮影年月日)"
		cRec[71] = "喀痰検査(塗抹鏡検 一般細菌)(所見)"
		cRec[72] = "喀痰検査(塗抹鏡検 抗酸菌)"
		cRec[73] = "喀痰検査(ガフキー号数)"
		cRec[74] = "便潜血"
		cRec[75] = "視力(裸眼右)"
		cRec[76] = "視力(矯正右)"
		cRec[77] = "視力(裸眼左)"
		cRec[78] = "視力(矯正左)"
		cRec[79] = "聴力(右1000Hz)"
		cRec[80] = "聴力(右4000Hz)"
		cRec[81] = "聴力(左1000Hz)"
		cRec[82] = "聴力(左4000Hz)"
		cRec[83] = "聴力(その他の所見)"
		cRec[84] = "眼底検査(キースワグナー分類)"
		cRec[85] = "眼底検査(シェイエ分類:H)"
		cRec[86] = "眼底検査(シェイエ分類:S)"
		cRec[87] = "眼底検査(SCOTT分類)"
		cRec[88] = "眼底検査(その他の所見)"
		cRec[89] = "眼底検査(実施理由)"
		cRec[90] = "その他の法定特殊健康診断"
		cRec[91] = "その他の法定検査"
		cRec[92] = "その他の検査"
		cRec[93] = "追加項目1"
		cRec[94] = "追加項目2"
		cRec[95] = "追加項目3"
		cRec[96] = "追加項目4"
		cRec[97] = "追加項目5"
		cRec[98] = "追加項目6"
		cRec[99] = "追加項目7"
		cRec[100] = "追加項目8"
		cRec[101] = "追加項目9"
		cRec[102] = "追加項目10"
		cRec[103] = "BMI判定"
		cRec[104] = "内臓脂肪面積判定"
		cRec[105] = "腹囲判定"
		cRec[106] = "血圧判定"
		cRec[107] = "総コレステロール判定"
		cRec[108] = "中性脂肪判定"
		cRec[109] = "HDLコレステロール判定"
		cRec[110] = "LDLコレステロール判定"
		cRec[111] = "GOT(AST)判定"
		cRec[112] = "GPT(ALT)判定"
		cRec[113] = "γ-GT(γ-GTP)判定"
		cRec[114] = "血清クレアチニン判定"
		cRec[115] = "血清尿酸判定"
		cRec[116] = "空腹時血糖判定"
		cRec[117] = "随時血糖判定"
		cRec[118] = "HbA1c判定"
		cRec[119] = "HbA1c（NGSP)判定"
		cRec[120] = "尿糖判定"
		cRec[121] = "尿蛋白判定"
		cRec[122] = "尿潜血判定"
		cRec[123] = "尿素窒素判定"
		cRec[124] = "尿ウロビリノーゲン判定"
		cRec[125] = "ヘマトクリット値判定"
		cRec[126] = "血色素量(ヘモグロビン値)判定"
		cRec[127] = "赤血球数判定"
		cRec[128] = "白血球数判定"
		cRec[129] = "血小板数判定"
		cRec[130] = "視力(右)判定"
		cRec[131] = "視力(左)判定"
		cRec[132] = "追加項目判定1"
		cRec[133] = "追加項目判定2"
		cRec[134] = "追加項目判定3"
		cRec[135] = "追加項目判定4"
		cRec[136] = "追加項目判定5"
		cRec[137] = "追加項目判定6"
		cRec[138] = "追加項目判定7"
		cRec[139] = "追加項目判定8"
		cRec[140] = "追加項目判定9"
		cRec[141] = "追加項目判定10"
		cRec[142] = "コメント"
		cRec[143] = "総合判定"
		cRec[144] = "受診勧奨区分"
		cRec[145] = "指導状態"
		cRec[146] = "再検査区分"
		cRec[147] = "一次健診日"
		cRec[148] = "結果通知区分"
		cRec[149] = "メタボリック判定(血圧リスク)"
		cRec[150] = "メタボリック判定(血糖リスク)"
		cRec[151] = "メタボリック判定(脂質リスク)"
		cRec[152] = "メタボリック判定(リスクカウント)"
		cRec[153] = "支援レベル(血圧リスク)"
		cRec[154] = "支援レベル(血糖リスク)"
		cRec[155] = "支援レベル(脂質リスク)"
		cRec[156] = "支援レベル(喫煙リスク)"
		cRec[157] = "支援レベル(リスクカウント)"
		cRec[158] = "メタボリックシンドローム判定"
		cRec[159] = "支援レベル"
		cRec[160] = "医師の診断(判定)"
		cRec[161] = "健康診断を実施した医師の氏名"
		cRec[162] = "医師の意見"
		cRec[163] = "意見を述べた医師の氏名"
		cRec[164] = "歯科医師による健康診断"
		cRec[165] = "歯科医師による健康診断を実施した歯科医師の氏名"
		cRec[166] = "歯科医師の意見"
		cRec[167] = "意見を述べた歯科医師の氏名"
		cRec[168] = "備考"
		cRec[169] = "服薬１_血圧"
		cRec[170] = "血圧_薬剤"
		cRec[171] = "血圧_服薬理由"
		cRec[172] = "服薬２_血糖"
		cRec[173] = "血糖_薬剤"
		cRec[174] = "血糖_服薬理由"
		cRec[175] = "服薬３_脂質"
		cRec[176] = "脂質_薬剤"
		cRec[177] = "脂質_服薬理由"
		cRec[178] = "既往歴１_脳血管"
		cRec[179] = "既往歴２_心血管"
		cRec[180] = "既往歴３_腎不全人工透析"
		cRec[181] = "貧血"
		cRec[182] = "喫煙"
		cRec[183] = "２０歳からの体重変化"
		cRec[184] = "３０分以上の運動習慣"
		cRec[185] = "歩行又は身体活動"
		cRec[186] = "歩行速度"
		cRec[187] = "１年間の体重変化"
		cRec[188] = "食べ方１_早食い等"
		cRec[189] = "食べ方２_就寝前"
		cRec[190] = "食べ方３_夜食間食"
		cRec[191] = "食習慣"
		cRec[192] = "飲酒"
		cRec[193] = "飲酒量"
		cRec[194] = "睡眠"
		cRec[195] = "生活習慣の改善"
		cRec[196] = "保健指導の希望"
		cRec[197] = "報告対象区分"
		cRec[198] = "保健指導からの除外"
		cRec[199] = "取込年月日"
		cRec[200] = "胸部X線判定"
		cRec[201] = "胸部判定アルファベット"
		cRec[202] = "心電図判定アルファベット"
		cRec[203] = "胸部レントゲン検査"
		cRec[204] = "胸部レントゲン判定"
		cRec[205] = "尿糖"
		cRec[206] = "尿蛋白"
		cRec[207] = "聴力(右1000Hz)"
		cRec[208] = "聴力(右4000Hz)"
		cRec[209] = "聴力(左1000Hz)"
		cRec[210] = "聴力(左4000Hz)"
		cRec[211] = "心電図検査"
		cRec[212] = "心電図判定"
		//writer.Write(cRec)
		for c, cell = range cRec {
			sheet.Cell(2, c).Value = cell
		}

		// 4行目移行（データ）
		r = 3
		inRecsMax := len(inRecs)
		for J := 1; J < inRecsMax; J++ {
			for I, _ = range cRec {
				cRec[I] = ""
			}

			if inRecs[J][4] == coRec[0] {
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

				// 9.加入番号

				// 10.扶養番号

				// 11.受診者区分

				// 12.性別
				cRec[12] = inRecs[J][8]

				// 13.氏名漢字
				cRec[13] = inRecs[J][9]

				// 14.氏名カナ
				cRec[14] = string(norm.NFKC.Bytes([]byte(inRecs[J][10])))

				// 15.生年月日
				cRec[15] = WaToSeireki(inRecs[J][11])

				// 16.実施年度
				cRec[16] = nendo(inRecs[J][15])

				// 17.年齢
				cRec[17] = inRecs[J][12]

				// 18.受診日
				cRec[18] = strings.Replace(inRecs[J][15], "-", "/", -1)

				// 19.健診区分
				cRec[19] = "事業者健診"

				// 20.医療機関コード
				cRec[20] = "013-61"

				// 21.医療機関名称
				cRec[21] = "医療法人社団　松英会"

				// 22.機関コード
				cRec[22] = "1311131242"

				// 23.機関名称
				cRec[23] = "医療法人社団　松英会"

				// 24.機関住所
				cRec[24] = "143-0027 大田区中馬込1-5-8"

				// 25.受付NO
				cRec[25] = inRecs[J][16]

				// 26.身長
				cRec[26] = inRecs[J][17]

				// 27.体重
				cRec[27] = inRecs[J][18]

				// 28.BMI
				cRec[28] = inRecs[J][19]

				// 29.内臓脂肪面積

				// 30.腹囲
				cRec[30] = inRecs[J][20]

				// 31.業務歴

				// 32.既往歴
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

				cRec[32] = kiou

				// 33.自覚症状
				cRec[33] = syoken(inRecs[J][41] + " " + inRecs[J][42] + " " + inRecs[J][43])

				// 34.他覚症状
				cRec[34] = syoken(inRecs[J][44] + " " + inRecs[J][45] + " " + inRecs[J][46])

				// 35.収縮期血圧(その他)

				// 36.収縮期血圧(２回目)
				cRec[36] = inRecs[J][47]

				// 37.収縮期血圧(１回目)
				cRec[37] = inRecs[J][48]

				// 38.拡張期血圧(その他)

				// 39.拡張期血圧(２回目)
				cRec[39] = inRecs[J][49]

				// 40.拡張期血圧(１回目)
				cRec[40] = inRecs[J][50]

				// 41.採血時間

				// 42.総コレステロール
				cRec[42] = inRecs[J][53]

				// 43.中性脂肪
				cRec[43] = inRecs[J][54]

				// 44.HDLコレステロール
				cRec[44] = inRecs[J][55]

				// 45.LDLコレステロール
				cRec[45] = inRecs[J][56]

				// 46.GOT(AST)
				cRec[46] = inRecs[J][57]

				// 47.GPT(ALT)
				cRec[47] = inRecs[J][58]

				// 48.γ-GT(γ-GTP)
				cRec[48] = inRecs[J][59]

				// 49.血清クレアチニン
				cRec[49] = inRecs[J][60]

				// 50.血清尿酸
				cRec[50] = inRecs[J][61]

				// 51.空腹時血糖
				// 52.随時血糖
				if syokugo(inRecs[J][51], inRecs[J][52]) {
					cRec[52] = inRecs[J][62]
				} else {
					cRec[51] = inRecs[J][62]
				}

				// 53.HbA1c
				// 54.HbA1c(NGSP)
				cRec[54] = inRecs[J][63]

				// 55.尿糖
				cRec[55] = inRecs[J][64]

				// 56.尿蛋白
				cRec[56] = inRecs[J][65]

				// 57.尿潜血
				cRec[57] = inRecs[J][66]

				// 58.尿素窒素
				cRec[58] = inRecs[J][67]

				// 59.尿ウロビリノーゲン
				cRec[59] = inRecs[J][68]

				// 60.ヘマトクリット値
				cRec[60] = inRecs[J][69]

				// 61.血色素量(ヘモグロビン値)
				cRec[61] = inRecs[J][70]

				// 62.赤血球数
				cRec[62] = inRecs[J][71]

				// 63.貧血検査実施理由

				// 64.白血球数
				cRec[64] = inRecs[J][72]

				// 65.血小板数
				cRec[65] = inRecs[J][73]

				// 66.血清アミラーゼ
				cRec[66] = inRecs[J][74]

				// 67.心電図(所見)
				cRec[67] = syoken(inRecs[J][76] + " " + inRecs[J][77] + " " + inRecs[J][78] + " " + inRecs[J][79])

				// 68.心電図(実施理由)

				// 69.胸部X線検査(所見)
				cRec[69] = syoken(inRecs[J][81] + " " + inRecs[J][82] + " " + inRecs[J][83])

				// 70.胸部X線検査(撮影年月日)
				if inRecs[J][80] != "" {
					cRec[70] = strings.Replace(inRecs[J][15], "-", "/", -1)
				}

				// 71.喀痰検査(塗抹鏡検 一般細菌)(所見)

				// 72.喀痰検査(塗抹鏡検 抗酸菌)

				// 73.喀痰検査(ガフキー号数)

				// 74.便潜血
				cRec[74] = inRecs[J][84]

				// 75.視力(裸眼右)
				cRec[75] = eye(inRecs[J][85])

				// 76.視力(矯正右)
				cRec[76] = eye(inRecs[J][86])

				// 77.視力(裸眼左)
				cRec[77] = eye(inRecs[J][87])

				// 78.視力(矯正左)
				cRec[78] = eye(inRecs[J][88])

				// 79.聴力(右1000Hz)
				cRec[79] = syokenumu(inRecs[J][97])

				// 80.聴力(右4000Hz)
				cRec[80] = syokenumu4k(inRecs[J][99], inRecs[J][101])

				// 81.聴力(左1000Hz)
				cRec[81] = syokenumu(inRecs[J][98])

				// 82.聴力(左4000Hz)
				cRec[82] = syokenumu4k(inRecs[J][100], inRecs[J][102])

				// 83.聴力(その他の所見)

				// 84.眼底検査(キースワグナー分類)

				// 85.眼底検査(シェイエ分類:H)

				// 86.眼底検査(シェイエ分類:S)

				// 87.眼底検査(SCOTT分類)

				// 88.眼底検査(その他の所見)

				// 89.眼底検査(実施理由)

				// 90.その他の法定特殊健康診断

				// 91.その他の法定検査

				// 92.その他の検査

				// 93.追加項目1

				// 94.追加項目2

				// 95.追加項目3

				// 96.追加項目4

				// 97.追加項目5

				// 98.追加項目6

				// 99.追加項目7

				// 100.追加項目8

				// 101.追加項目9

				// 102.追加項目10

				// 103.BMI判定
				cRec[103] = inRecs[J][103]

				// 104.内臓脂肪面積判定

				// 105.腹囲判定
				cRec[105] = inRecs[J][104]

				// 106.血圧判定
				cRec[106] = string(norm.NFKC.Bytes([]byte(inRecs[J][105])))

				// 107.総コレステロール判定
				cRec[107] = inRecs[J][106]

				// 108.中性脂肪判定
				cRec[108] = inRecs[J][107]

				// 109.HDLコレステロール判定
				cRec[109] = inRecs[J][108]

				// 110.LDLコレステロール判定
				cRec[110] = inRecs[J][109]

				// 111.GOT(AST)判定
				cRec[111] = inRecs[J][110]

				// 112.GPT(ALT)判定
				cRec[112] = inRecs[J][111]

				// 113.γ-GT(γ-GTP)判定
				cRec[113] = inRecs[J][112]

				// 114.血清クレアチニン判定
				cRec[114] = inRecs[J][113]

				// 115.血清尿酸判定
				cRec[115] = inRecs[J][114]

				// 116.空腹時血糖判定
				// 117.随時血糖判定
				if syokugo(inRecs[J][51], inRecs[J][52]) {
					cRec[117] = toH(inRecs[J][62])
				} else {
					cRec[116] = inRecs[J][115]
				}

				// 118.HbA1c判定

				// 119.HbA1c（NGSP)判定
				cRec[119] = inRecs[J][116]

				// 120.尿糖判定
				cRec[120] = inRecs[J][117]

				// 121.尿蛋白判定
				cRec[121] = inRecs[J][118]

				// 122.尿潜血判定
				cRec[122] = inRecs[J][119]

				// 123.尿素窒素判定
				cRec[123] = inRecs[J][120]

				// 124.尿ウロビリノーゲン判定
				cRec[124] = inRecs[J][121]

				// 125.ヘマトクリット値判定
				cRec[125] = inRecs[J][122]

				// 126.血色素量(ヘモグロビン値)判定
				cRec[126] = inRecs[J][123]

				// 127.赤血球数判定
				cRec[127] = inRecs[J][124]

				// 128.白血球数判定
				cRec[128] = inRecs[J][125]

				// 129.血小板数判定
				cRec[129] = inRecs[J][126]

				// 130.視力(右)判定
				cRec[130] = eyeHantei(inRecs[J][127], inRecs[J][128])

				// 131.視力(左)判定
				cRec[131] = eyeHantei(inRecs[J][129], inRecs[J][130])

				// 132.追加項目判定1

				// 133.追加項目判定2

				// 134.追加項目判定3

				// 135.追加項目判定4

				// 136.追加項目判定5

				// 137.追加項目判定6

				// 138.追加項目判定7

				// 139.追加項目判定8

				// 140.追加項目判定9

				// 141.追加項目判定10

				// 142.コメント

				// 143.総合判定
				cRec[143] = inRecs[J][131]

				// 144.受診勧奨区分

				// 145.指導状態

				// 146.再検査区分

				// 147.一次健診日

				// 148.結果通知区分

				// 149.メタボリック判定(血圧リスク)

				// 150.メタボリック判定(血糖リスク)

				// 151.メタボリック判定(脂質リスク)

				// 152.メタボリック判定(リスクカウント)

				// 153.支援レベル(血圧リスク)

				// 154.支援レベル(血糖リスク)

				// 155.支援レベル(脂質リスク)

				// 156.支援レベル(喫煙リスク)

				// 157.支援レベル(リスクカウント)

				// 158.メタボリックシンドローム判定
				cRec[158] = inRecs[J][133]

				// 159.支援レベル
				cRec[159] = inRecs[J][134]

				// 160.医師の診断(判定)
				cRec[160] = inRecs[J][132]

				// 161.健康診断を実施した医師の氏名
				cRec[161] = "寺門　節雄"

				// 162.医師の意見

				// 163.意見を述べた医師の氏名

				// 164.歯科医師による健康診断

				// 165.歯科医師による健康診断を実施した歯科医師の氏名

				// 166.歯科医師の意見

				// 167.意見を述べた歯科医師の氏名

				// 168.備考

				// 169.服薬１_血圧
				cRec[169] = inRecs[J][135]

				// 170.血圧_薬剤

				// 171.血圧_服薬理由

				// 172.服薬２_血糖
				cRec[172] = inRecs[J][136]

				// 173.血糖_薬剤

				// 174.血糖_服薬理由

				// 175.服薬３_脂質
				cRec[175] = inRecs[J][137]

				// 176.脂質_薬剤

				// 177.脂質_服薬理由

				// 178.既往歴１_脳血管
				cRec[178] = inRecs[J][138]

				// 179.既往歴２_心血管
				cRec[179] = inRecs[J][139]

				// 180.既往歴３_腎不全人工透析
				cRec[180] = inRecs[J][140]

				// 181.貧血
				cRec[181] = inRecs[J][141]

				// 182.喫煙
				cRec[182] = inRecs[J][142]

				// 183.２０歳からの体重変化
				cRec[183] = inRecs[J][143]

				// 184.３０分以上の運動習慣
				cRec[184] = inRecs[J][144]

				// 185.歩行又は身体活動
				cRec[185] = inRecs[J][145]

				// 186.歩行速度
				cRec[186] = inRecs[J][146]

				// 187.１年間の体重変化
				cRec[187] = inRecs[J][147]

				// 188.食べ方１_早食い等
				cRec[188] = inRecs[J][148]

				// 189.食べ方２_就寝前
				cRec[189] = inRecs[J][149]

				// 190.食べ方３_夜食間食
				cRec[190] = inRecs[J][150]

				// 191.食習慣
				cRec[191] = inRecs[J][151]

				// 192.飲酒
				cRec[192] = inRecs[J][152]

				// 193.飲酒量
				cRec[193] = inRecs[J][153]

				// 194.睡眠
				cRec[194] = inRecs[J][154]

				// 195.生活習慣の改善
				cRec[195] = inRecs[J][155]

				// 196.保健指導の希望
				cRec[196] = inRecs[J][156]

				// 197.報告対象区分

				// 198.保健指導からの除外

				// 199.取込年月日

				// 200.胸部X線判定
				cRec[200] = string(norm.NFKC.Bytes([]byte(inRecs[J][80])))

				// 201.胸部判定アルファベット
				cRec[201] = string(norm.NFKC.Bytes([]byte(inRecs[J][80])))

				// 202.心電図判定アルファベット
				cRec[202] = string(norm.NFKC.Bytes([]byte(inRecs[J][75])))

				// 203.胸部レントゲン検査
				cRec[203] = hanteiCode(string(norm.NFKC.Bytes([]byte(inRecs[J][80]))))

				// 204.胸部レントゲン判定
				cRec[204] = hanteiCode(string(norm.NFKC.Bytes([]byte(inRecs[J][80]))))

				// 205.尿糖
				cRec[205] = nyou(inRecs[J][64])

				// 206.尿蛋白
				cRec[206] = nyou(inRecs[J][65])

				// 207.聴力(右1000Hz)
				cRec[207] = syokenumuCode(syokenumu(inRecs[J][97]))

				// 208.聴力(右4000Hz)
				cRec[208] = syokenumuCode(syokenumu4k(inRecs[J][99], inRecs[J][101]))

				// 209.聴力(左1000Hz)
				cRec[209] = syokenumuCode(syokenumu(inRecs[J][98]))

				// 210.聴力(左4000Hz)
				cRec[210] = syokenumuCode(syokenumu4k(inRecs[J][100], inRecs[J][102]))

				// 211.心電図検査
				cRec[211] = hanteiCode(string(norm.NFKC.Bytes([]byte(inRecs[J][75]))))

				// 212.心電図判定
				cRec[212] = hanteiCode(string(norm.NFKC.Bytes([]byte(inRecs[J][75]))))

				//writer.Write(cRec)
				for c, cell = range cRec {
					sheet.Cell(r, c).Value = cell
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
			if coRec[0] == inRec[4] || inRec[4] == "所属cd１" {
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
				for c, cell := range jRec {
					sheet.Cell(r, c).Value = cell
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
