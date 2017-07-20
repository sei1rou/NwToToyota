package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"golang.org/x/text/encoding/japanese"
	"golang.org/x/text/transform"
	"golang.org/x/text/unicode/norm"
	"io"
	"log"
	"os"
	"strconv"
	"strings"
)

func failOnError(err error) {
	if err != nil {
		log.Fatal("Error:", err)
	}
}

func main() {
	flag.Parse()

	//ログファイル準備
	logfile, err := os.OpenFile("./log.txt", os.O_APPEND|os.O_CREATE|os.O_WRONLY, os.ModePerm)
	failOnError(err)
	defer logfile.Close()

	log.SetOutput(logfile)

	//入力ファイル準備
	infile, err := os.Open(flag.Arg(0))
	failOnError(err)
	defer infile.Close()

	//書き込みファイル準備
	outfile, err := os.Create("./沖電気カスタマアドテック健診データ（健保用）.csv")
	failOnError(err)
	defer outfile.Close()

	reader := csv.NewReader(transform.NewReader(infile, japanese.ShiftJIS.NewDecoder()))
	reader.Comma = '\t'
	writer := csv.NewWriter(transform.NewWriter(outfile, japanese.ShiftJIS.NewEncoder()))
	writer.UseCRLF = true

	log.Print("Start\r\n")
	//タイトル行を読み出す
	_, err = reader.Read() // 1行読み出す
	if err != io.EOF {
		failOnError(err)
	}

	for {
		record, err := reader.Read() // 1行読み出す
		if err == io.EOF {
			break
		} else {
			failOnError(err)
		}

		var out_record []string
		errPersonalInfo := record[18] + "," + record[8]

		//  1:実施区分
		out_record = append(out_record, "1")

		//  2:プログラム種別
		out_record = append(out_record, "030")

		//  3:実施年月日
		out_record = append(out_record, strings.Replace(record[17], "-", "", -1))

		//  4:健診機関番号
		out_record = append(out_record, "1311131242")

		//  5:健診機関名称
		out_record = append(out_record, "医療法人社団　松英会　馬込中央診療所")

		//  6:健診機関郵便番号
		out_record = append(out_record, "143-0027")

		//  7:健診機関所在地
		out_record = append(out_record, "東京都大田区中馬込１－５－８")

		//  8:健診機関電話番号
		out_record = append(out_record, "03-3773-6773")

		//  9:保険者番号
		if record[4] == "" {
			log.Print("保健者番号なし:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, record[4])

		//  10:被保険者等記号
		if record[5] == "" {
			log.Print("保険証記号なし:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, record[5])

		//  11:被保険者等番号
		if record[6] == "" {
			log.Print("保険証番号なし:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, record[6])

		//  12:カナ氏名
		if record[7] == "" {
			log.Print("カナ氏名なし:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, string(norm.NFKC.Bytes([]byte(record[7]))))

		//  13:漢字氏名
		if record[8] == "" {
			log.Print("漢字氏名なし:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, record[8])

		//  14:生年月日
		if record[10] == "" {
			log.Print("生年月日なし:" + errPersonalInfo + "\r\n")
		}
		if WaToSeireki(record[10]) == "err" {
			log.Print("生年月日エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, WaToSeireki(record[10]))

		//  15:男女区分
		if record[9] == "" {
			log.Print("性別なし:" + errPersonalInfo + "\r\n")
		}
		if Sei(record[9]) == "err" {
			log.Print("性別エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, Sei(record[9]))

		//  16:郵便番号
		if record[11] == "" {
			log.Print("郵便番号なし:", errPersonalInfo+"\r\n")
		}
		out_record = append(out_record, record[11])

		//  17:住所
		if record[12] == "" {
			log.Print("住所なし:", errPersonalInfo+"\r\n")
		}
		out_record = append(out_record, strings.Trim(record[12]+"　"+record[13], "　"))

		//  18:受診券整理番号
		out_record = append(out_record, "")

		//  19:受診券有効期限
		out_record = append(out_record, "")

		//  20:健診種別コード
		out_record = append(out_record, "")

		//  21:事業所コード
		out_record = append(out_record, "")

		//  22:社員番号
		out_record = append(out_record, record[14])

		//  23:予備
		out_record = append(out_record, "")

		//  24:予備
		out_record = append(out_record, "")

		//  25:予備
		out_record = append(out_record, "")

		//  26:予備
		out_record = append(out_record, "")

		//  27:予備
		out_record = append(out_record, "")

		//  28:予備
		out_record = append(out_record, "")

		//  29:身長
		if record[19] == "" {
			log.Print("身長なし:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, record[19])

		//  30:体重
		if record[20] == "" {
			log.Print("体重なし:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, record[20])

		//  31:BMI
		if record[21] == "" {
			log.Print("BMIなし:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, record[21])

		//  32:腹囲測定法
		//  33:腹囲
		if record[22] == "" {
			out_record = append(out_record, "")
			out_record = append(out_record, "")
		} else {
			out_record = append(out_record, "1")
			out_record = append(out_record, record[22])
		}

		//  34:内臓脂肪面積
		out_record = append(out_record, "")

		//  35:下限値
		out_record = append(out_record, "")

		//  36:上限値
		out_record = append(out_record, "")

		//  37:収縮期血圧区分
		//  38:収縮期血圧
		if record[23] == "" {
			out_record = append(out_record, "")
			out_record = append(out_record, "")
		} else if record[25] == "" {
			out_record = append(out_record, "2")
			out_record = append(out_record, record[23])
		} else {
			d38_1, _ := strconv.Atoi(record[23])
			d38_2, _ := strconv.Atoi(record[25])
			d38 := (d38_1 + d38_2) / 2
			out_record = append(out_record, "1")
			out_record = append(out_record, fmt.Sprint(d38))
		}

		//  39:下限値
		out_record = append(out_record, "")

		//  40:上限値
		out_record = append(out_record, "")

		//  41:欠番
		out_record = append(out_record, "")

		//  42:欠番
		out_record = append(out_record, "")

		//  43:欠番
		out_record = append(out_record, "")

		//  44:拡張期血圧区分
		//  45:拡張期血圧
		if record[24] == "" {
			out_record = append(out_record, "")
			out_record = append(out_record, "")
		} else if record[26] == "" {
			out_record = append(out_record, "2")
			out_record = append(out_record, record[24])
		} else {
			d45_1, _ := strconv.Atoi(record[24])
			d45_2, _ := strconv.Atoi(record[26])
			d45 := (d45_1 + d45_2) / 2
			out_record = append(out_record, "1")
			out_record = append(out_record, fmt.Sprint(d45))
		}

		//  46:下限値
		out_record = append(out_record, "")

		//  47:上限値
		out_record = append(out_record, "")

		//  48:欠番
		out_record = append(out_record, "")

		//  49:欠番
		out_record = append(out_record, "")

		//  50:欠番
		out_record = append(out_record, "")

		//  51:総コレステロール測定法
		//  52:総コレステロール
		if record[27] == "" {
			out_record = append(out_record, "")
			out_record = append(out_record, "")
		} else {
			out_record = append(out_record, "01")
			out_record = append(out_record, record[27])
		}

		//  53:下限値
		out_record = append(out_record, "")

		//  54:上限値
		out_record = append(out_record, "")

		//  55:HDLコレステロール測定法
		//  56:HDLコレステロール
		if record[28] == "" {
			log.Print("HDL-Cなし:" + errPersonalInfo + "\r\n")
			out_record = append(out_record, "")
			out_record = append(out_record, "")
		} else {
			out_record = append(out_record, "01")
			out_record = append(out_record, record[28])
		}

		//  57:下限値
		out_record = append(out_record, "")

		//  58:上限値
		out_record = append(out_record, "")

		//  59:LDLコレステロール測定法
		//  60:LDLコレステロール
		if record[29] == "" {
			log.Print("LDL-Cなし:" + errPersonalInfo + "\r\n")
			out_record = append(out_record, "")
			out_record = append(out_record, "")
		} else {
			out_record = append(out_record, "01")
			out_record = append(out_record, record[29])
		}

		//  61:下限値
		out_record = append(out_record, "")

		//  62:上限値
		out_record = append(out_record, "")

		//  63:中性脂肪測定法
		//  64:中性脂肪
		if record[30] == "" {
			log.Print("中性脂肪なし:" + errPersonalInfo + "\r\n")
			out_record = append(out_record, "")
			out_record = append(out_record, "")
		} else {
			out_record = append(out_record, "01")
			out_record = append(out_record, record[30])
		}

		//  65:下限値
		out_record = append(out_record, "")

		//  66:上限値
		out_record = append(out_record, "")

		//  67:GOT(AST)測定法
		//  68:GOT(AST)
		if record[31] == "" {
			log.Print("GOTなし:" + errPersonalInfo + "\r\n")
			out_record = append(out_record, "")
			out_record = append(out_record, "")
		} else {
			out_record = append(out_record, "02")
			out_record = append(out_record, record[31])
		}

		//  69:下限値
		out_record = append(out_record, "")

		//  70:上限値
		out_record = append(out_record, "")

		//  71:GPT(ALT)測定法
		//  72:GPT(ALT)
		if record[32] == "" {
			log.Print("GPTなし:" + errPersonalInfo + "\r\n")
			out_record = append(out_record, "")
			out_record = append(out_record, "")
		} else {
			out_record = append(out_record, "02")
			out_record = append(out_record, record[32])
		}

		//  73:下限値
		out_record = append(out_record, "")

		//  74:上限値
		out_record = append(out_record, "")

		//  75:γGTP測定法
		//  76:γGTP
		if record[33] == "" {
			log.Print("γGTP:" + errPersonalInfo + "\r\n")
			out_record = append(out_record, "")
			out_record = append(out_record, "")
		} else {
			out_record = append(out_record, "01")
			out_record = append(out_record, record[33])
		}
		//  77:下限値
		out_record = append(out_record, "")

		//  78:上限値
		out_record = append(out_record, "")

		//  79:空腹時血糖測定法
		//  80:空腹時血糖
		//  81:下限値
		//  82:上限値
		//  83:随時血糖測定法
		//  84:随時血糖
		//  85:下限値
		//  86:上限値
		//  87:HbA1c測定法
		//  88:HbA1c
		//  89:下限値
		//  90:上限値
		Eattime, _ := strconv.Atoi(record[108])
		if record[34] == "" {
			if record[35] == "" {
				log.Print("血糖・HbA1cなし:" + errPersonalInfo + "\r\n")
				out_record = append(out_record, "") // 79
				out_record = append(out_record, "") // 80
				out_record = append(out_record, "") // 81
				out_record = append(out_record, "") // 82
				out_record = append(out_record, "") // 83
				out_record = append(out_record, "") // 84
				out_record = append(out_record, "") // 85
				out_record = append(out_record, "") // 86
				out_record = append(out_record, "") // 87
				out_record = append(out_record, "") // 88
				out_record = append(out_record, "") // 89
				out_record = append(out_record, "") // 90
			} else {
				out_record = append(out_record, "")         // 79
				out_record = append(out_record, "")         // 80
				out_record = append(out_record, "")         // 81
				out_record = append(out_record, "")         // 82
				out_record = append(out_record, "")         // 83
				out_record = append(out_record, "")         // 84
				out_record = append(out_record, "")         // 85
				out_record = append(out_record, "")         // 86
				out_record = append(out_record, "14")       // 87
				out_record = append(out_record, record[35]) // 88
				out_record = append(out_record, "")         // 89
				out_record = append(out_record, "")         // 90
			}
		} else {
			if (record[107] == "とった") && (Eattime < 10) {
				out_record = append(out_record, "")         // 79
				out_record = append(out_record, "")         // 80
				out_record = append(out_record, "")         // 81
				out_record = append(out_record, "")         // 82
				out_record = append(out_record, "01")       // 83
				out_record = append(out_record, record[34]) // 84
				out_record = append(out_record, "")         // 85
				out_record = append(out_record, "")         // 86
				if record[35] == "" {
					log.Print("空腹時血糖検査なし:" + errPersonalInfo + "\r\n")
					out_record = append(out_record, "") // 87
					out_record = append(out_record, "") // 88
					out_record = append(out_record, "") // 89
					out_record = append(out_record, "") // 90
				} else {
					out_record = append(out_record, "14")       // 87
					out_record = append(out_record, record[35]) // 88
					out_record = append(out_record, "")         // 89
					out_record = append(out_record, "")         // 90
				}
			} else {
				out_record = append(out_record, "01")       // 79
				out_record = append(out_record, record[34]) // 80
				out_record = append(out_record, "")         // 81
				out_record = append(out_record, "")         // 82
				out_record = append(out_record, "")         // 83
				out_record = append(out_record, "")         // 84
				out_record = append(out_record, "")         // 85
				out_record = append(out_record, "")         // 86
				if record[35] == "" {
					out_record = append(out_record, "") // 87
					out_record = append(out_record, "") // 88
					out_record = append(out_record, "") // 89
					out_record = append(out_record, "") // 90
				} else {
					out_record = append(out_record, "14")       // 87
					out_record = append(out_record, record[35]) // 88
					out_record = append(out_record, "")         // 89
					out_record = append(out_record, "")         // 90
				}
			}
		}

		//  91:赤血球
		out_record = append(out_record, record[36])

		//  92:下限値
		out_record = append(out_record, "")

		//  93:上限値
		out_record = append(out_record, "")

		//  94:血色素量
		out_record = append(out_record, record[37])

		//  95:下限値
		out_record = append(out_record, "")

		//  96:上限値
		out_record = append(out_record, "")

		//  97:ヘマトクリット
		out_record = append(out_record, record[38])

		//  98:下限値
		out_record = append(out_record, "")

		//  99:上限値
		out_record = append(out_record, "")

		//  100:貧血検査実施理由
		out_record = append(out_record, "")

		//  101:MCHC
		out_record = append(out_record, record[39])

		//  102:下限値
		out_record = append(out_record, "")

		//  103:上限値
		out_record = append(out_record, "")

		//  104:尿糖定性区分
		//  105:尿糖定性
		if record[40] == "" {
			out_record = append(out_record, "")
			out_record = append(out_record, "")
		} else {
			if Nyou(record[40]) == "err" {
				log.Print("尿糖コード変換エラー:" + errPersonalInfo + "\r\n")
			}
			out_record = append(out_record, "1")
			out_record = append(out_record, Nyou(record[40]))
		}

		//  106:尿蛋白定性区分
		//  107:尿蛋白定性
		if record[41] == "" {
			out_record = append(out_record, "")
			out_record = append(out_record, "")
		} else {
			if Nyou(record[41]) == "err" {
				log.Print("尿蛋白コード変換エラー:" + errPersonalInfo + "\r\n")
			}
			out_record = append(out_record, "1")
			out_record = append(out_record, Nyou(record[41]))
		}

		//  108:心電図（所見の有無）
		//  109:心電図所見
		if record[42] == "" {
			out_record = append(out_record, "")
			out_record = append(out_record, "")
		} else {
			if Syokenumu(record[42]) == "err" {
				log.Print("心電図判定変換エラー:" + errPersonalInfo + "\r\n")
			}
			out_record = append(out_record, Syokenumu(record[42]))
			out_record = append(out_record, strings.Trim(record[43]+"　"+record[44]+"　"+record[45]+"　"+record[46], "　"))
		}

		//  110:心電図実施理由
		out_record = append(out_record, "")

		//	※沖カスタマアドテックは眼底検査なし。便宜上右のデータを入れる事にする。
		//  111:眼底検査（シェイエ分類：Ｈ）
		if HyScConv(record[49]) == "err" {
			log.Print("Hy変換エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, HyScConv(record[49]))

		//  112:欠番
		out_record = append(out_record, "")

		//  113:欠番
		out_record = append(out_record, "")

		//  114:眼底検査（シェイエ分類：Ｓ）
		if HyScConv(record[48]) == "err" {
			log.Print("Sc変換エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, HyScConv(record[48]))

		//  115:欠番
		out_record = append(out_record, "")

		//  116:欠番
		out_record = append(out_record, "")

		//  117:眼底検査（キースワグナー分類）
		if KwConv(record[47]) == "err" {
			log.Print("Kw変換エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, KwConv(record[47]))

		//  118:欠番
		out_record = append(out_record, "")

		//  119:欠番
		out_record = append(out_record, "")

		//  120:眼底検査（SCOTT分類)
		if ScottConv(record[50]) == "err" {
			log.Print("Scott変換エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, ScottConv(record[50]))

		//  121:眼底検査(その他の所見)
		if record[55] == "" {
			out_record = append(out_record, "")
		} else {
			out_record = append(out_record, strings.Trim(record[55]+record[56]+record[57]+record[58], " "))
		}

		//  122:眼底検査実施理由
		out_record = append(out_record, "")

		//  123:他覚症状
		//  124:他覚症状所見
		if record[59] == "" {
			out_record = append(out_record, "")
			out_record = append(out_record, "")
		} else {
			if Syokenumu(record[59]) == "err" {
				log.Print("内科診察判定変換エラー:" + errPersonalInfo + "\r\n")
			}
			out_record = append(out_record, Syokenumu(record[59]))
			out_record = append(out_record, strings.Trim(record[60]+"　"+record[61]+"　"+record[62], "　"))
		}

		//  125:具体的な既往歴
		D125_1 := KiouJoin(record[63], record[64], record[65])
		D125_2 := KiouJoin(record[66], record[67], record[68])
		D125_3 := KiouJoin(record[69], record[70], record[71])
		D125_4 := KiouJoin(record[72], record[73], record[74])
		D125_5 := KiouJoin(record[75], record[76], record[77])
		D125 := strings.Trim(D125_1+"／"+D125_2+"／"+D125_3+"／"+D125_4+"／"+D125_5, "／")
		out_record = append(out_record, D125)

		//  126:欠番
		out_record = append(out_record, "")

		//  127:欠番
		out_record = append(out_record, "")

		//  128:欠番
		out_record = append(out_record, "")

		//  129:欠番
		out_record = append(out_record, "")

		//  130:欠番
		out_record = append(out_record, "")

		//  131:欠番
		out_record = append(out_record, "")

		//  132:欠番
		out_record = append(out_record, "")

		//  133:欠番
		out_record = append(out_record, "")

		//  134:欠番
		out_record = append(out_record, "")

		//  135:既往歴
		if D125 == "" {
			out_record = append(out_record, "2")
		} else {
			out_record = append(out_record, "1")
		}

		//  136:保健指導レベル
		if HokenConv(record[78]) == "err" {
			log.Print("保健指導レベルエラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, HokenConv(record[78]))

		//  137:医師の診断（判定）
		if record[79] == "" {
			log.Print("総合判定が入っていません:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, record[79])

		//  138:健診実施医師名
		out_record = append(out_record, "寺門　節雄")

		//  139:メタボリックシンドローム判定
		if MetaboConv(record[81]) == "err" {
			log.Print("メタボリックシンドローム判定エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, MetaboConv(record[81]))

		//  140:病歴(脳血管疾患）
		if YesNo(record[82]) == "err" {
			log.Print("病歴(脳血管疾患)エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, YesNo(record[82]))

		//  141:病歴（心血管）
		if YesNo(record[83]) == "err" {
			log.Print("病歴(心血管)エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, YesNo(record[83]))

		//  142:病歴（腎不全・人工透析）
		if YesNo(record[84]) == "err" {
			log.Print("病歴(腎不全・人工透析)エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, YesNo(record[84]))

		//  143:病歴（貧血）
		if YesNo(record[85]) == "err" {
			log.Print("病歴(貧血)エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, YesNo(record[85]))

		//  144:服薬1（血圧）
		if YesNo(record[86]) == "err" {
			log.Print("服薬1(血圧)エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, YesNo(record[86]))

		//  145:服薬2（血糖）
		if YesNo(record[87]) == "err" {
			log.Print("服薬2(血糖)エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, YesNo(record[87]))

		//  146:服薬3（脂質）
		if YesNo(record[88]) == "err" {
			log.Print("服薬3(脂質)エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, YesNo(record[88]))

		//  147:喫煙区分
		if YesNo(record[89]) == "err" {
			log.Print("喫煙区分エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, YesNo(record[89]))

		//  148:欠番
		out_record = append(out_record, "")

		//  149:問　20歳から10kg以上の体重増
		if YesNo(record[90]) == "err" {
			log.Print("20歳から体重増加エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, YesNo(record[90]))

		//  150:問　30分以上の運動習慣
		if YesNo(record[91]) == "err" {
			log.Print("運動習慣:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, YesNo(record[91]))

		//  151:問　身体活動を1日1時間以上
		if YesNo(record[92]) == "err" {
			log.Print("身体活動エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, YesNo(record[92]))

		//  152:問　歩行速度 同性同年齢比較で速い
		if YesNo(record[93]) == "err" {
			log.Print("歩行速度エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, YesNo(record[93]))

		//  153:問　1年間の体重変化±3kg以上
		if YesNo(record[94]) == "err" {
			log.Print("1年間の体重変化エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, YesNo(record[94]))

		//  154:問　食べ方（早食い）
		if Eat(record[95]) == "err" {
			log.Print("早食いエラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, Eat(record[95]))

		//  155:問　就寝前2H以内夕食、3回/週
		if YesNo(record[96]) == "err" {
			log.Print("就寝前夕食エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, YesNo(record[96]))

		//  156:問　食べ方（夜食/間食）3回/週
		if YesNo(record[97]) == "err" {
			log.Print("夜食エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, YesNo(record[97]))

		//  157:問　朝食抜き3回/週
		if YesNo(record[98]) == "err" {
			log.Print("朝食抜きエラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, YesNo(record[98]))

		//  158:問　飲酒習慣
		if Sake(record[99]) == "err" {
			log.Print("飲酒エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, Sake(record[99]))

		//  159:飲酒量（飲酒日）
		if Sakeryo(record[100]) == "err" {
			log.Print("飲酒量/日エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, Sakeryo(record[100]))

		//  160:問　睡眠で休養がとれる
		if YesNo(record[101]) == "err" {
			log.Print("睡眠休養エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, YesNo(record[101]))

		//  161:生活習慣改善意識
		if Seikatsu(record[102]) == "err" {
			log.Print("生活習慣改善エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, Seikatsu(record[102]))

		//  162:保健指導利用希望
		if YesNo(record[103]) == "err" {
			log.Print("保健指導希望エラー:" + errPersonalInfo + "\r\n")
		}
		out_record = append(out_record, YesNo(record[103]))

		//  163:自覚症状
		//  164:自覚症状所見
		D164 := strings.Trim(record[104]+"　"+record[105]+"　"+record[106], "　")
		if D164 == "" {
			out_record = append(out_record, "2")
		} else if strings.Index(D164, "特になし") >= 0 {
			out_record = append(out_record, "2")
		} else {
			out_record = append(out_record, "1")
		}

		out_record = append(out_record, D164)

		//  165:服薬1(血圧)(薬剤名)
		out_record = append(out_record, "")

		//  166:服薬1(血圧)(実施理由)
		out_record = append(out_record, "")

		//  167:服薬2(血糖)(薬剤名)
		out_record = append(out_record, "")

		//  168:服薬2(血糖)(実施理由)
		out_record = append(out_record, "")

		//  169:服薬3(脂質)(薬剤名)
		out_record = append(out_record, "")

		//  170:服薬3(脂質)(実施理由)
		out_record = append(out_record, "")

		//  171:採血時間(食後)
		if (record[107] == "とった") && (Eattime < 10) {
			out_record = append(out_record, "1")
		} else {
			out_record = append(out_record, "2")
		}

		// １行書き出す
		writer.Write(out_record)
	}
	writer.Flush()
	log.Print("Finesh !\r\n")

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
			return fmt.Sprint(yi) + m + d
		}
	}
}

func Sei(s string) string {

	switch s {
	case "男":
		s = "1"
	case "女":
		s = "2"
	case "性別":
		s = "性別"
	default:
		s = "err"
	}
	return s
}

func Nyou(s string) string {

	switch s {
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
		s = "5"
	case "5+":
		s = "5"
	default:
		s = "err"
	}
	return s
}

func Syokenumu(s string) string {

	switch s {
	case "Ａ":
		s = "2"
	case "Ｂ":
		s = "1"
	case "Ｃ":
		s = "1"
	case "Ｄ":
		s = "1"
	case "Ｅ":
		s = "1"
	case "Ｆ":
		s = "1"
	case "Ｇ":
		s = "1"
	default:
		s = "err"
	}
	return s
}

func KwConv(s string) string {

	switch s {
	case "":
		s = ""
	case "０":
		s = "1"
	case "Ⅰ":
		s = "2"
	case "Ⅱ":
		s = "3"
	case "Ⅲ":
		s = "5"
	case "Ⅳ":
		s = "6"
	case "Ⅴ":
		s = "6"
	case "Ⅱａ":
		s = "3"
	case "Ⅱｂ":
		s = "4"
	case "Ⅲａ":
		s = "5"
	case "Ⅲｂ":
		s = "5"
	case "Ⅰａ":
		s = "2"
	case "Ⅰｂ":
		s = "2"
	default:
		s = "err"
	}
	return s
}

func HyScConv(s string) string {

	switch s {
	case "":
		s = ""
	case "０":
		s = "1"
	case "１":
		s = "2"
	case "２":
		s = "3"
	case "３":
		s = "4"
	case "４":
		s = "5"
	default:
		s = "err"
	}
	return s
}

func ScottConv(s string) string {

	switch s {
	case "":
		s = ""
	case "０":
		s = ""
	case "Ⅰ":
		s = "1"
	case "Ⅱ":
		s = "3"
	case "Ⅲ":
		s = "4"
	case "Ⅳ":
		s = "6"
	case "Ⅴ":
		s = "7"
	case "Ⅵ":
		s = "9"
	case "Ⅱａ":
		s = "3"
	case "Ⅱｂ":
		s = "3"
	case "Ⅲａ":
		s = "4"
	case "Ⅲｂ":
		s = "5"
	case "Ⅰａ":
		s = "1"
	case "Ⅰｂ":
		s = "2"
	default:
		s = "err"
	}
	return s
}

func KiouJoin(b, a, t string) string {

	s := ""

	if a != "" {
		a = a + "才"
	}

	if b != "" {
		s = strings.Trim(strings.Replace(b+" "+a+" "+t, "  ", " ", -1), " ")
	}

	return s
}

func HokenConv(s string) string {

	switch s {
	case "":
		s = ""
	case "積極的支援レベル":
		s = "1"
	case "動機づけ支援レベル":
		s = "2"
	case "情報提供レベル":
		s = "3"
	case "判定不能":
		s = "4"
	default:
		s = "err"
	}

	return s
}

func MetaboConv(s string) string {

	switch s {
	case "":
		s = ""
	case "基準該当":
		s = "1"
	case "予備群該当":
		s = "2"
	case "非該当":
		s = "3"
	case "判定不能":
		s = "4"
	default:
		s = "err"
	}
	return s
}

func YesNo(s string) string {

	switch s {
	case "":
		s = ""
	case "はい":
		s = "1"
	case "いいえ":
		s = "2"
	default:
		s = "err"
	}
	return s
}

func Eat(s string) string {

	switch s {
	case "":
		s = ""
	case "速い":
		s = "1"
	case "普通":
		s = "2"
	case "遅い":
		s = "3"
	default:
		s = "err"
	}
	return s
}

func Sake(s string) string {

	switch s {
	case "":
		s = ""
	case "毎日":
		s = "1"
	case "時々":
		s = "2"
	case "飲まない":
		s = "3"
	default:
		s = "err"
	}
	return s
}

func Sakeryo(s string) string {

	switch s {
	case "":
		s = ""
	case "１合未満":
		s = "1"
	case "１～２合未満":
		s = "2"
	case "２～３合未満":
		s = "3"
	case "３合以上":
		s = "4"
	default:
		s = "err"
	}
	return s
}

func Seikatsu(s string) string {

	switch s {
	case "":
		s = ""
	case "しない":
		s = "1"
	case "思う":
		s = "2"
	case "始めた":
		s = "3"
	case "６ヶ月経過":
		s = "4"
	case "６ヶ月以上":
		s = "5"
	default:
		s = "err"
	}
	return s
}
