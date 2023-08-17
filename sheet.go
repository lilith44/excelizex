package excelizex

import (
	"errors"
	"fmt"
	"github.com/cyclonevox/excelizex/extra"
	"github.com/cyclonevox/excelizex/style"
	"github.com/xuri/excelize/v2"
	"reflect"
	"regexp"
	"strconv"
	"strings"
)

type sheet struct {
	// 表名
	name string
	// 格式化后的数据
	// 顶栏提示
	notice string
	// 表头
	header []string
	// 不渲染的列号
	omitCols map[int]struct{}
	// 数据
	data [][]any
	// 下拉选项 暂时只支持单列
	pd *pullDown

	// 分布和style分配语句的配置
	// v:part -> k:style string
	styleRef map[string][]*style.Parsed
	// 写入到第几行,主要用于标记生成excel中的表时，需要续写的位置
	writeRow int
}

func NewSheet(sheetName string, a any, omitColNames, redColNames []string) *sheet {
	if sheetName == "" {
		panic("sheet cannot be empty")
	}

	s := &sheet{
		name:     sheetName,
		omitCols: make(map[int]struct{}),
		styleRef: make(map[string][]*style.Parsed),
		writeRow: 0,
	}
	if a != nil {
		s.initSheetData(a, omitColNames, redColNames)
	}

	return s
}

func (s *sheet) Excel() *File {
	if s.name == "" {
		panic("need a sheet name at least")
	}

	return New().AddFormattedSheets(s)
}

func (s *sheet) initSheetData(a any, omitColNames, redColNames []string) {
	typ := reflect.TypeOf(a)
	val := reflect.ValueOf(a)

	if typ.Kind() == reflect.Ptr {
		typ = typ.Elem()
	}

	// 如果作为Slice类型的传入对象，则还需要注意拆分后进行处理
	switch typ.Kind() {
	case reflect.Slice:
		for i := 0; i < val.Len(); i++ {
			if i == 0 {
				s.setHeaderByStruct(val.Index(i).Interface(), omitColNames, redColNames)
			}

			s.data = append(s.data, getRowData(val.Index(i).Interface()))
		}
	case reflect.Struct:
		s.setHeaderByStruct(a, omitColNames, redColNames)
	}

	return
}

// SetHeaderByStruct 方法会检测结构体中的excel标签，以获取结构体表头
func (s *sheet) setHeaderByStruct(a any, omitColNames, redColNames []string) *sheet {
	typ := reflect.TypeOf(a)
	val := reflect.ValueOf(a)
	if typ.Kind() == reflect.Ptr {
		typ = typ.Elem()
		val = val.Elem()
	}

	if typ.Kind() != reflect.Struct {
		panic(errors.New("generate function support using struct only"))
	}

	omitColMapping := make(map[string]struct{})
	for _, name := range omitColNames {
		omitColMapping[name] = struct{}{}
	}

	redColMapping := make(map[string]struct{})
	for _, name := range redColNames {
		redColMapping[name] = struct{}{}
	}

	for i := 0; i < typ.NumField(); i++ {
		typeField := typ.Field(i)

		partTag := typeField.Tag.Get("excel")
		if partTag == "" {
			continue
		} else {

			// 判断是excel tag 是指向哪个部分
			params := strings.Split(partTag, "|")
			if len(params) > 0 {
				switch extra.Part(params[0]) {
				case extra.NoticePart:
					s.notice = val.Field(i).String()

					// 添加提示样式映射
					styleString := typeField.Tag.Get("style")
					if styleString == "" {
						continue
					}
					_noticeStyle := style.TagParse(styleString).Parse()
					_noticeStyle.Cell.StartCell = style.Cell{Col: "A", Row: 1}
					_noticeStyle.Cell.EndCell = style.Cell{Col: "A", Row: 1}
					s.styleRef[fmt.Sprintf("%s", extra.NoticePart)] = []*style.Parsed{&_noticeStyle}

				case extra.HeaderPart:
					// todo： 现在header的style暂时不能交叉设置，原因是会被覆盖，需要在后续改动
					s.header = append(s.header, params[1])
					if _, ok := omitColMapping[params[1]]; ok {
						s.omitCols[len(s.header)-1] = struct{}{}

						continue
					}

					styleString := typeField.Tag.Get("style")
					if _, ok := redColMapping[params[1]]; ok {
						styleString = "default-header-red"
					}

					if styleString == "" {
						continue
					}

					colName, err := excelize.ColumnNumberToName(len(s.header) - len(s.omitCols))
					if err != nil {
						panic(err)
					}
					headerStyle := style.TagParse(styleString).Parse()

					pp, ok := s.styleRef[string(extra.HeaderPart)]
					if !ok || !reflect.DeepEqual(headerStyle.StyleNames, pp[len(pp)-1].StyleNames) {
						headerStyle.Cell.StartCell = style.Cell{Col: colName, Row: 2}
						headerStyle.Cell.EndCell = style.Cell{Col: colName, Row: 2}

						s.styleRef[string(extra.HeaderPart)] = append(s.styleRef[string(extra.HeaderPart)], &headerStyle)
					} else {
						pp[len(pp)-1].Cell.EndCell.Col = colName
					}

					// todo :暂不支持 太累了抱歉
					// styleString = typeField.Tag.Get("data-style")
					//dataStyle := style.TagParse(styleString).Parse(extra.DataPart)
					//s.styleRef[fmt.Sprintf("%s-%s", extra.DataPart, params[1])] = dataStyle
				}
			}

		}
	}

	return s
}

func getRowData(row any) (list []any) {
	typ := reflect.TypeOf(row)
	val := reflect.ValueOf(row)

	if typ.Kind() == reflect.Ptr {
		typ = typ.Elem()
		val = val.Elem()
	}

	switch typ.Kind() {
	case reflect.Struct:
		for j := 0; j < typ.NumField(); j++ {
			field := typ.Field(j)

			hasTag := field.Tag.Get("excel")
			if hasTag != "" && hasTag != "notice" {
				list = append(list, val.Field(j).Interface())
			}
		}
	case reflect.Slice:
		for i := 0; i < val.Len(); i++ {
			list = append(list, val.Index(i).Interface())
		}

	default:
		panic("support struct only")
	}

	return
}

// findHeaderColumnName 寻找表头名称或者是列名称
func (s *sheet) findHeaderColumnName(headOrColName string) (columnName string, err error) {
	for i, h := range s.header {
		if h == headOrColName {
			columnName, err = excelize.ColumnNumberToName(i + 1)

			return
		}
	}

	regular := `[A-Z]+`
	reg := regexp.MustCompile(regular)
	if !reg.MatchString(headOrColName) {
		panic("plz use A-Z ColName or HeaderName for option name ")
	}

	columnName = headOrColName

	return
}

// SetOptions 设置下拉的选项
func (s *sheet) SetOptions(headOrColName string, options any) *sheet {
	name, err := s.findHeaderColumnName(headOrColName)
	if err != nil {
		panic(err)
	}

	pd := newPullDown().addOptions(name, options)

	if s.pd == nil {
		s.pd = pd
	} else {
		s.pd.merge(pd)
	}

	return s
}

// nextWriteRow 会获取目前该写入的行
// 每次调用该方法表示行数增长 返回 A1 A2... 等名称
func (s *sheet) nextWriteRow(num ...int) string {
	if len(num) > 0 {
		s.writeRow += num[0]
	} else {
		s.writeRow++
	}

	return "A" + strconv.FormatInt(int64(s.writeRow), 10)
}

func (s *sheet) getWriteRow() string {
	return "A" + strconv.FormatInt(int64(s.writeRow), 10)
}

func (s *sheet) resetWriteRow() string {
	s.writeRow = 1

	return "A" + strconv.FormatInt(int64(s.writeRow), 10)
}
