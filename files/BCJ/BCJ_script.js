//=============================
// 0. 사전 정의
//=============================
//함수 : DOM 선택자
function $(parameter, startNode) {
    if (!startNode)
        return document.querySelector(parameter)
    else
        return startNode.querySelector(parameter)
}
function $$(parameter, startNode) {
    if (!startNode)
        return document.querySelectorAll(parameter)
    else
        return startNode.querySelectorAll(parameter)
}
//함수 : 일시정지(화면 전환용)
let sleep = async (ms) => {
    return new Promise((r) => setTimeout(r, ms))
}

//허용 확장자 : xls, xlsx
let allowedExtensionsReg = /(\.xls|\.xlsx)$/i
let allowedExtensionsArr = ["xls","xslx"]

//업로드된 엑셀 파일 묶움
let excelArr = []//엑셀 데이터 저장소
let falsedArr = []//검토용 데이터는 여기에 집어넣기
let mergedArr = []//합산용 데이터는 여기에 집어넣기
let monthArr = []//연월 모음집(정렬 후)
const headerList = [/*엑셀 행렬 서식*/
    "번호","도서관","대출자번호","신청자","ISBN","연도/가격","서명","신청일","연월","상태","월 건수","중복신청","검토사유"
]
const referList = [//입력 헤더 참조 : [0] => [1] 로 변경
    ["제목","서명"],
    ["도서명","서명"],
    ["정가","가격"],
    ["금액","가격"],
    ["이름","신청자"],
    ["진행상태","상태"]
]
const ignoreStateList = [//제외하는 상태
    "취소함","도서관거절","신청거절","신청취소","재고없음"
]
//검토 옵션
let checkYearPrice = false//발행년 & 가격 검토 여부
let noCheckdongne = true//동네서점 초과 신청 검토 제외 여부

//=============================
// 1. 파일 업로드(드래그 방식)
//=============================
//A. 드래그 방식
let dropZone = document.getElementById('dropZone')

function allowDrag(e) {
    if (true) {  // Test that the item being dragged is a valid one
        e.dataTransfer.dropEffect = 'move'
        e.preventDefault()
    }
}

// A-1. 드래그 시작
window.addEventListener('dragenter', (e) => {
    dropZone.classList.add("show")
})

// A-2. 드래그 중
dropZone.addEventListener('dragenter', allowDrag)
dropZone.addEventListener('dragover', allowDrag)

// A-3. 드래그 나감
dropZone.addEventListener('dragleave', (e) => {
    dropZone.classList.remove("show")
})

// A-4. 드래그 놓음 : 업로드
dropZone.addEventListener('drop', (e) => {
    dropZone.classList.remove("show")
    uploadFiles(e)
})

//=============================
// 2. 파일 업로드
//=============================
let uploadFiles = async (e) => {
    e.preventDefault()

    let uploadArr = e.dataTransfer.files
        //업로드 파일이 없으면 중단
        if (!uploadArr) return
    let readArr = []
    //엑셀 파일 읽기 함수
    let readFile = (file) => {
        return new Promise(resolve => {
            let reader = new FileReader()
            reader.onload = (e) => {
                resolve({
                    name:file.name.replace(/\.[^/.]+$/, ""),
                    type:/(?:\.([^.]+))?$/.exec(file.name)[1],
                    binary:e.target.result
                })
            }
            reader.readAsBinaryString(file)
        })
    }
    //읽기 함수 병렬실행 모음집 만들기
    for(let i=0;i<uploadArr.length;i++) {
        let file = uploadArr[i]
        //허용하지 않는 확장자는 건너뛰기
        if(!allowedExtensionsReg.exec(file.name)){
            continue
        } else {
            readArr.push(readFile(file))
        }
    }
    //올바른 확장자 파일이 없으면 종료
    if (readArr.length <= 0) {
        alert("* " + uploadArr.length.toString() + "개 파일 중 [확장자가 올바른 파일]이 없습니다.")
        return
    } else {
        //발행년 & 가격 검토 여부 확인
        checkYearPrice = confirm("* 발행년 및 가격을 검토하겠습니까?")
        //읽기 함수 병렬실행
        excelArr = []
        excelArr = await Promise.all(readArr)
        //엑셀 분석 시작
        handleFile()
    }
}

//=============================
// 2. 각 엑셀 파일 정리
//=============================
let handleFile = async () => {
    //준비 페이지 잠시 보여주기
    $("#pageBody1").style.display = "none"
    $("#pageBody2").style.display = "block"
    $("#pageBody2").innerHTML = "자료 입력 중..."
    await sleep(10)

    //입력된 엑셀 분석
    excelArr.forEach(lib => {
        lib.data = parseExcel(lib)
    })

    //체크박스 활성화
    $("#checkYearPrice").disabled = false
    $("#checkYearPrice_desc").classList.add("active")
    if (checkYearPrice === true) {
        $("#checkYearPrice").checked = true
    } else {
        $("#checkYearPrice").checked = false
    }
    $("#noCheckdongne").disabled = false
    $("#noCheckdongne_desc").classList.add("active")

    //분석 시작
    $("#pageBody2").innerHTML = "자료 분석 중..."
    analyseExcel()
}

//=============================
// 3. 체크박스 변경 따른 신규 분석
//=============================
//연도/가격 체크 사항이 변경되면 분석을 다시 실시
$("#checkYearPrice").addEventListener("change", () => {
    //연도/가격 검토여부 변경
    checkYearPrice = $("#checkYearPrice").checked
    //재분석 시작
    $("#pageBody2").innerHTML = "자료 다시 분석 중..."
    analyseExcel()
})
//동네서점 초과 검토 체크 사항이 변경되면 분석을 다시 실시
$("#noCheckdongne").addEventListener("change", () => {
    //연도/가격 검토여부 변경
    noCheckdongne = $("#noCheckdongne").checked
    //재분석 시작
    $("#pageBody2").innerHTML = "자료 다시 분석 중..."
    analyseExcel()
})

//=============================
// 4. 분석 및 출력
//=============================
let analyseExcel = async () => {

    //준비 페이지 잠시 보여주기
    $("#pageBody1").style.display = "none"
    $("#pageBody2").style.display = "block"
    await sleep(10)

    //0. 데이터 입력 준비
    mergedArr = []//합산용 데이터
    falsedArr = []//검토용 데이터

    //1. 전관 필요데이터 합치기(취소 자료 제외)
    monthArr = []
    excelArr.forEach(lib => {
        //비치희망, 동네서점 구분 (도서관명 추가 방식)
        if (lib.name.indexOf("동네") < 0) {
            //비치희망 루프
            lib.data.forEach((line,I) => {
                //열 정보 참조 수정
                referList.forEach(refer => {
                    if (line[refer[0]] !== undefined)
                        line[refer[1]] = line[refer[0]]
                })
                //취소된 데이터는 무시
                if (ignoreStateList.indexOf(line["상태"]) >= 0) return
                //도서관명 추가 (비치희망 : <파일명>)
                line["도서관"] = lib.name
                //연도/가격 추가
                line["연도/가격"] = line["발행년"] + "<hr>" + line["가격"].replace(/\B(?=(\d{3})+(?!\d))/g, ",")
                //연월(신청년+신청월) 추가
                let time = new Date(line["신청일"])
                line["연월"] = time.getFullYear().toString() + "." + ("0" + (time.getMonth() + 1).toString()).slice(-2)
                if (monthArr.indexOf(line["연월"]) < 0) {
                    monthArr.push(line["연월"])
                    monthArr.sort().reverse()
                }
                mergedArr.push(line)
            })
        } else {
            //동네서점 루프
            lib.data.forEach(line => {
                //열 정보 참조 수정
                referList.forEach(refer => {
                    if (line[refer[0]] !== undefined)
                        line[refer[1]] = line[refer[0]]
                })
                //취소된 데이터는 무시
                if (ignoreStateList.indexOf(line["상태"]) >= 0) return
                //도서관명 추가 (동네서점 : "#동네_" + <도서관명>)
                line["도서관"] = "#동네_" + line["도서관명"].replace(/양산시립|도서관/g,"").trim()
                //연도/가격 추가
                line["연도/가격"] = line["발행년"] + "<hr>" + line["가격"].replace(/\B(?=(\d{3})+(?!\d))/g, ",")
                //연월(신청년+신청월) 추가
                let time = new Date(line["신청일"])
                line["연월"] = time.getFullYear().toString() + "." + ("0" + (time.getMonth() + 1).toString()).slice(-2)
                if (monthArr.indexOf(line["연월"]) < 0) {
                    monthArr.push(line["연월"])
                    monthArr.sort().reverse()
                }
                mergedArr.push(line)
            })
        }
    })

    //2. 자료 정렬(연월 내림, 이용자 오름, 신청일 내림)
    mergedArr.sort((a,b) => {
        let KeyA = [a["연월"],b["연월"]]
        let KeyB = [a["신청자"],b["신청자"]]
        let KeyC = [a["신청일"],b["신청일"]]

        if (KeyA[0] > KeyA[1]) {
            return -1
        } else if (KeyA[0] < KeyA[1]) {
            return 1
        } else {
            if (KeyB[0] < KeyB[1]) {
                return -1
            } else if (KeyB[0] > KeyB[1]) {
                return 1
            } else {
                if (KeyC[0] > KeyC[1]) {
                    return -1
                } else if (KeyC[0] < KeyC[1]) {
                    return 1
                } else {
                    return 0
                }
            }
        }
    })

    //3. 월 건수, 중복신청, 연도+가격 분석
    mergedArr.forEach((line1,i) => {
        line1.error = []
        let monthCount = 0
        let duplicate = 0
        //월 건수
        for (let j = i;j < mergedArr.length;j++) {
            let line2 = mergedArr[j]
            if (line1["대출자번호"] === line2["대출자번호"] &&
                line1["연월"] === line2["연월"] &&
                    //양쪽 다 동네서점이 아닌 경우에만 초과 검토 (체크 시)
                    (noCheckdongne === false || 
                    (noCheckdongne === true && line1["도서관"].indexOf("동네") < 0 && line2["도서관"].indexOf("동네") < 0 )))
                    monthCount += 1
        }
        //월 건수 3회 초과 시 오류 기록해두기
        if (monthCount > 3) {
            if (line1.error.indexOf("월 건수") < 0) line1.error.push("월 건수")
        }
        //중복신청
        for (let j = 0;j < mergedArr.length;j++) {
            let line3 = mergedArr[j]
            if (line1["ISBN"].length > 0 &&
                i !== j && line1["대출자번호"] === line3["대출자번호"] &&
                line1["ISBN"] === line3["ISBN"]) {
                    duplicate += 1
            }
            //오류 기록해두기
            if (duplicate > 0) {
                if (line1.error.indexOf("중복신청") < 0) line1.error.push("중복신청")
            }
        } 
        //발행년 & 가격(검토 대상일 때)
        if (checkYearPrice === true) {
            if (parseInt(line1["발행년"]) < parseInt(line1["연월"].substr(0,4)) - 5 ||
                parseInt(line1["가격"]) >= 50000) {
                    //오류 기록해두기
                    if (line1.error.indexOf("연도/가격") < 0) line1.error.push("연도/가격")
            }
        }
        line1["월 건수"] = monthCount.toString()
        line1["중복신청"] = duplicate.toString()
        if (line1.error.length > 0) {
            line1["검토사유"] = line1.error.join(",")
            let falsedId = line1["대출자번호"] + "_" + line1["연월"]
            if (falsedArr.indexOf(falsedId) < 0) falsedArr.push(falsedId)
        } else {
            line1["검토사유"] = ""
        }
    })

    //4. 검토 대상이 아닌 데이터 삭제
    /*
    mergedArr.forEach((line,i) => {
        if (falsedArr.indexOf(line["대출자번호"] + "_" + line["연월"]) < 0) {
            mergedArr.splice(i,1)
        }
    })
    */
    let deleteIndex = mergedArr.length - 1
    while (deleteIndex >= 0) {
        if (falsedArr.indexOf(mergedArr[deleteIndex]["대출자번호"] + "_" + mergedArr[deleteIndex]["연월"]) < 0) {
            mergedArr.splice(deleteIndex,1)
        }
      
        deleteIndex -= 1;
    }

    //5. 자료 출력
    let table = document.getElementById("dataTable")
    //기존 데이터 비우기
    while (table.hasChildNodes()) {
        table.removeChild(table.lastChild);
    }
    //a. 헤더 추가
    let header = table.insertRow(-1)
    headerList.forEach(col => {
        let newCol = document.createElement("th")
        newCol.innerText = col
        //헤더 명시
        newCol.classList.add("header")
        if (col === "서명") newCol.classList.add("title")
        header.appendChild(newCol)
    })
    //b. 데이터 행 추가
    mergedArr.forEach((line,i) => {
        let newRow = table.insertRow(-1)
        headerList.forEach((col) => {
            let newCol = newRow.insertCell()
            newCol.innerText = line[col]
            //바디 명시
            newCol.classList.add("body")
            //번호 매기기
            if (col === "번호") newCol.innerText = (i+1).toString()
            //제목 좌측 정렬
            if (col === "서명") newCol.classList.add("title")
            //연도/가격 : innerText가 아닌 innerHTML, 글자 작게
            if (col === "연도/가격") {
                newCol.innerHTML = line[col]
                newCol.classList.add("year_price")
            }
            //신청자가 달라짐 - 상단 테두리 굵은 검정색
            if (i > 0 && mergedArr[i]["신청자"] !== mergedArr[i-1]["신청자"]) {
                newCol.classList.add("user_changed")
            }
            //연월이 달라짐 - 상단 테두리 굵은 빨간색
            if (i > 0 && mergedArr[i]["연월"] !== mergedArr[i-1]["연월"]) {
                newCol.classList.remove("user_changed")
                newCol.classList.add("month_changed")
            }
            //최근 연월이 아니면 - 약간 투명하게
            if (mergedArr[i]["연월"] !== monthArr[0]) {
                newCol.classList.add("month_passed")
            }
            //검토 대상 행 - 주황색 배경
            if (mergedArr[i].error.length > 0) {
                newCol.classList.add("falsed")
                //검토 대상 셀(월 건수 초과 or 중복신청 초과 & 발행년/가격 문제) - 보라색 배경
                if (mergedArr[i].error.indexOf(col) >= 0) {
                    newCol.classList.remove("falsed")
                    newCol.classList.add("whyFalsed")
                }
            }
            //처리중 or 소장중 - 초록색 강조
            if (col === "상태" && (line[col] !== "신청" && line[col] !== "신청중")) {
                newCol.classList.add("processed")
            }
        })
    })
    //표, 표가 있는 페이지 보여주기
    $("#pageBody1").style.display = "block"
    $("#pageBody2").style.display = "none"
    document.getElementById("dataTable").classList.remove("hidden")

    //6. 엑셀 출력 버튼 준비
    let button = document.getElementById("writeExcel")

    //엑셀 파일 이름 준비
    let today = new Date()
    let todayText = today.getFullYear().toString() + "." + ("0" + (today.getMonth() + 1).toString()).slice(-2) + "." + ("0" + today.getDate().toString()).slice(-2) + "."
    let fileName = todayText + ' 비치희망 초과중복 분석.xlsx'
    if (checkYearPrice) {
        fileName = "(연도가격 검토) " + fileName
    } else {
        fileName = "(연도가격 미검토) " + fileName
    }

    //버튼을 누르면 엑셀 출력
    button.onclick = () => {
        writeExcel(document.getElementById("dataTable"),fileName,
            {legend://범례
                [
                    ["ffcc6e","검토 대상"],
                    ["7e067e","검토 사유"],
                    ["32de84","처리중 / 소장중 자료"]
                ],
            colWidth://열 너비
                [
                    6,//번호
                    9,//도서관
                    18,//대출자번호
                    12,//신청자
                    17,//ISBN
                    9,//연도/가격
                    50,//서명
                    22,//신청일
                    9,//연월
                    9,//상태
                    9,//월 건수
                    9,//중복신청
                    15
                ]
            }
        )
    }

    //출력 버튼 활성화
    button.disabled = false
}