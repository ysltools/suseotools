//함수 : RGB, RGBA를 HEX로 반환
const rgba2Excelhex = (rgba) => {
    let hex = `${rgba.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*(\d+\.{0,1}\d*))?\)$/).slice(1).map((n, i) => (i === 3 ? Math.round(parseFloat(n) * 255) : parseFloat(n)).toString(16).padStart(2, '0').replace('NaN', '')).join('')}`
    if (hex.length > 6) {
        hex = hex.substring(0,6)
    }
    
    return hex
}

//엑셀 첫줄 헤더 여부 확인
let checkFirstLineHeader = (input) => {
    let wb = XLSX.read(input.binary, {type: 'binary'})
    let arr = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1})
    
    if (arr[0].length > 1) {
        return true
    } else {
        return false
    }
}

//엑셀 입력 분석
let parseExcel = (input) => {
    let wb = XLSX.read(input.binary, {type: 'binary'})
    //1차 변환 : 헤더 찾기
    let findHeaderArr = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1})
    let headerRow = 0
    //헤더 정의 : 3개 이상의 셀이 있는 행이 처음 등장
    for (let i = 0;i < findHeaderArr.length;i++) {
        if (findHeaderArr[i].length > 2) {
            headerRow = i
            break
        }
    }
    //2차 변환 : 활용 데이터
    let obj = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{raw:false,range:headerRow})
    
    return obj
}

//엑셀 출력
let writeExcel = (tableDOM, outputName, opt) => {
    //## 1. 엑셀 파일 생성 ##
    //ArrayBuffer 만들어주는 함수
    function s2ab(s) {
        var buf = new ArrayBuffer(s.length) //convert s to arrayBuffer
        var view = new Uint8Array(buf)  //create uint8array as viewer
        for (var i=0; i<s.length; i++) view[i] = s.charCodeAt(i) & 0xFF //convert to octet
        return buf
    }
    //워크북 생성
    let wb = XLSX.utils.book_new()
    //시트 생성
    wb.SheetNames.push("export")


    //## 2. 엑셀 데이터 정리 ##
    //자료 저장소
    let sheetArr = []//시트 행 묶음 = 엑셀 전체 데이터
    let rowArr = []//각 행
    let colData = {}//각 셀
    let wscols = []//열 옵션(너비, 숨기기 등)

    //엑셀 출력 옵션
    if (opt !== undefined) {
        //범례: 첫줄에 범례 추가
        if (opt.legend !== undefined) {
            rowArr = []//행 초기화
            opt.legend.forEach(elArr => {
                colData = {}//셀 초기화
                    colData.v = ""
                    colData.t = "s"
                    colData.s = {fill:{fgColor:{rgb:elArr[0]}}}
                rowArr.push(colData)//셀 입력
                
                colData = {}//셀 초기화
                    colData.v = elArr[1]
                    colData.t = "s"
                    colData.s = {alignment:{vertical:"center",horizontal:"left"},font:{sz:8}}
                rowArr.push(colData)//셀 입력
            })
            sheetArr.push(rowArr)//행 입력
        }
        //열 너비
        if (opt.colWidth !== undefined) {
            opt.colWidth.forEach((w,i) => {
                if (wscols[i] === undefined) wscols[i] = {}
                wscols[i].wch = w
            })
        }
        //열 숨기기
        if (opt.hideCol !== undefined) {
            opt.hideCol.forEach((hPos) => {
                if (wscols[hPos] === undefined) wscols[hPos] = {}
                wscols[hPos].hidden = true
            })
        }
    }

    //테이블에 따라 엑셀 자료 정리
    let table = tableDOM
    table.childNodes[0].childNodes.forEach((row,i,arrR) => {
        rowArr = []//행 초기화
        row.childNodes.forEach((cell,j,arrC) => {
            colData = {}//셀 초기화
            let cellStyle = window.getComputedStyle(cell)
            if (cell.innerHTML.length > 0) {
                colData.v = cell.innerHTML.replaceAll("<hr>","\n").replaceAll("<br>","\n")
                colData.t = "s"
            } else {
                colData.t = "z"
            }
            colData.s = {}
            //정렬 및 줄바꿈
            colData.s.alignment = {}
                colData.s.alignment.vertical = "center"
                colData.s.alignment.horizontal = cellStyle.getPropertyValue("text-align")
                //white-space:nowrap 이 있거나 첫줄이면 줄바꿈을 하지 않음
                if (i === 0 || cellStyle.getPropertyValue("white-space") === "nowrap") {
                    colData.s.alignment.wrapText = false
                } else {
                    colData.s.alignment.wrapText = true
                }
            //배경 (흰색이 아닐 경우에만 지정)
            if (rgba2Excelhex(cellStyle.getPropertyValue("background-color")) !== "000000") {
                colData.s.fill = {}
                colData.s.fill.fgColor = {rgb:rgba2Excelhex(cellStyle.getPropertyValue("background-color"))}
            }
            //폰트
            colData.s.font = {}
                colData.s.font.name = "arial"
                colData.s.font.color = {rgb:rgba2Excelhex(cellStyle.getPropertyValue("color"))}
                if (colData.v !== undefined && colData.v.indexOf("\n") >= 0) {
                    colData.s.font.sz = 9
                } else {
                    colData.s.font.sz = 11
                }
                if (cellStyle.getPropertyValue("font-weight") > 400) {
                    colData.s.font.bold = true
                } else {
                    colData.s.font.bold = false
                }
            //테두리
            colData.s.border = {}
                let direction = ["top","left"]
                if (i === arrR.length - 1) direction.push("bottom")
                if (j === arrC.length - 1) direction.push("right")
                direction.forEach(d => {
                    colData.s.border[d] = {}
                    colData.s.border[d].color = {rgb:rgba2Excelhex(cellStyle.getPropertyValue("border-" + d + "-color"))}
                    if (parseInt(cellStyle.getPropertyValue("border-" + d + "-width").replace("px","")) > 2) {
                        colData.s.border[d].style = "thick"
                    } else {
                        colData.s.border[d].style = "thin"
                    }
                })
                //공란일 경우 : 주변 셀 테두리 부여
                if (colData.t === "z") {
                    //첫 행이 아닐 경우 : "상단 행 동일 열의 셀" 하단 테두리 활성화 (해당 셀이 공란이 아닐 시)
                    if (sheetArr.length > 0 && sheetArr[sheetArr.length - 1][j] !== undefined && sheetArr[sheetArr.length - 1][j].t !== "z") {
                        let upperColData = sheetArr[sheetArr.length - 1][j]
                        upperColData.s.border.bottom = {}
                        upperColData.s.border.bottom.color = {rgb:rgba2Excelhex(cellStyle.getPropertyValue("border-bottom-color"))}
                        if (parseInt(cellStyle.getPropertyValue("border-bottom-width").replace("px","")) > 2) {
                            upperColData.s.border.bottom.style = "thick"
                        } else {
                            upperColData.s.border.bottom.style = "thin"
                        }
                    }
                    //첫 열이 아닐 경우 : "왼쪽 셀" 우측 테두리 활성화 (해당 셀이 공란이 아닐 시)
                    if (rowArr.length > 0 && rowArr[rowArr.length - 1].t !== "z") {
                        let beforeColData = rowArr[rowArr.length - 1]
                        beforeColData.s.border.right = {}
                        beforeColData.s.border.right.color = {rgb:rgba2Excelhex(cellStyle.getPropertyValue("border-right-color"))}
                        if (parseInt(cellStyle.getPropertyValue("border-right-width").replace("px","")) > 2) {
                            beforeColData.s.border.right.style = "thick"
                        } else {
                            beforeColData.s.border.right.style = "thin"
                        }
                    }
                }
            rowArr.push(colData)//셀 입력
        })
        sheetArr.push(rowArr)//행 입력
    })


    //## 3. 엑셀 출력 ##
    //입력 자료를 기반으로 시트 데이터 생성
    let ws = XLSX.utils.aoa_to_sheet(sheetArr)
    //시트 입력
    wb.Sheets["export"] = ws
    //열 너비 (있으면) 적용
    if (wscols.length > 0) ws['!cols'] = wscols
    //엑셀 파일 최종 작성
    let wbout = XLSX.write(wb, {bookType:'xlsx',  type: 'binary'})
    //엑셀 파일 다운로드
    saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), outputName)
}