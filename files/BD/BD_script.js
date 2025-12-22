//=============================
// 0. 사전 정의
//=============================
//함수 : DOM 선택자
function $(parameter, startNode) {
    if (!startNode)
        return document.querySelector(parameter);
    else
        return startNode.querySelector(parameter);
}
function $$(parameter, startNode) {
    if (!startNode)
        return document.querySelectorAll(parameter);
    else
        return startNode.querySelectorAll(parameter);
}
//함수 : 딜레이
function sleep(ms) {
    return new Promise((r) => setTimeout(r, ms));
}


//허용 확장자 : xls, xlsx
let allowedExtensionsReg = /(\.xls|\.xlsx)$/i
let allowedExtensionsArr = ["xls","xslx"]

//데이터 관리 객체
let dataObj = {//각 단계별 자료 정리
    step1:{B:{},D:{}},//1단계 : 비치희망 및 동네서점 각 엑셀 정보 저장
    step2:{},//2단계 : 합쳐진 엑셀 정보 저장
    step3:{},//3단계 : ISBN 중복 삭제한 엑셀 정보 저장
    step4:{}//4단계 : 반입용 엑셀 & 검토용 엑셀 정보 저장
}
let inputTypeArr = ["B","D"]//비치희망, 동네서점 루프용(B : 비치희망, D : 동네서점)
const headerRow = {//헤더 행 번호
    B:3,//비치희망 : 3행
    D:1//동네서점 : 2행
}
const headerList = [//엑셀 헤더
    "번호","구분","ISBN","검토","서명","저자","발행자","발행년","가격","상태","대출자번호","신청자","신청일"
]
const headerNotInputList = [//엑셀 헤더 중 입력 자료에는 없는 것 (차후 생성되는 헤더)
    "번호","구분","검토"
]
const referList = [//입력 헤더 참조 : [0] => [1] 로 변경
    ["제목","서명"],
    ["도서명","서명"],
    ["출판사","발행자"],
    ["정가","가격"],
    ["금액","가격"],
    ["이름","신청자"],
    ["진행상태","상태"]
]
const importHeaderList = [//반입용 엑셀 전용 헤더
    "서명","저자","발행자","발행년","낱권ISBN","낱권ISBN부가기호","가격","세트ISBN","세트ISBN부가기호",
    "KDC 분류기호","DDC 분류기호","기타 분류기호","청구기호_분류","청구기호_도서기호","청구기호_권책기호",
    "부서명","대등서명","공저자","판사항","편/권차","편제","발행지","면장수","물리적특성","크기","딸림자료",
    "총서표제","총서편차","수상주기","등록번호","복본기호","별치기호","URI"
]
const referImportList = {//반입용 엑셀 입력 헤더 참조 : value => key 로 변경
    "서명":"서명",
    "저자":"저자",
    "발행자":"발행자",
    "발행년":"발행년",
    "낱권ISBN":"ISBN",
    "가격":"가격"
}
const ignoreStateList = [//각 단계별로 제외하는 상태
    ["처리중","소장","취소함","도서관거절","신청거절","신청취소","입고","재고없음"],//ISBN, 서명 중복 조회 전 제외 상태
    ["도서관승인","재고확인중","승인","승인취소","대출","반납"]//ISBN, 서명 중복 조회 후 제외 상태
]

//엑셀 업로드 제어
let allowInput = true


//=============================
// 1. 파일 업로드
//=============================
//A. 드래그 방식
let dropZone = document.getElementById('dropZone')

function allowDrag(e) {
    if (true) {  // Test that the item being dragged is a valid one
        if (allowInput === true) {
            e.dataTransfer.dropEffect = 'move'
        } else {
            e.dataTransfer.dropEffect = 'none'
        }
        e.preventDefault()
    }
}

// A-1. 드래그 시작
window.addEventListener('dragenter', (e) => {
    if (allowInput === true) {
        dropZone.innerHTML = "여기에 엑셀 파일(들)을 올리세요"
        dropZone.classList.add("show1")
    } else {
        dropZone.innerHTML = "지금은 엑셀 파일을 올릴 수 없습니다"
        dropZone.classList.add("show2")
    }
})

// A-2. 드래그 중
dropZone.addEventListener('dragenter', allowDrag)
dropZone.addEventListener('dragover', allowDrag)

// A-3. 드래그 나감
dropZone.addEventListener('dragleave', (e) => {
    dropZone.classList.remove("show1","show2")
})

// A-4. 드래그 놓음 : 업로드(허용 시)
dropZone.addEventListener('drop', (e) => {
    dropZone.classList.remove("show1","sho2")
    if (allowInput === true) {
        uploadFiles(e)
    } else {
        e.preventDefault()
        return false
    }
})

//파일 업로드 헤더 체크 준비
inputTypeArr.forEach((type) => {
    //맨 앞에 명시
    let elDesc = document.createElement("div")
        elDesc.classList.add("checkExcelInputEl_desc")
        elDesc.innerHTML = "확인된 정보 :"
    $("#checkExcelInput" + type).appendChild(elDesc)
    //각 헤더별 체크
    headerList.forEach((head) => {
        let el = document.createElement("div")
            el.id = "checkExcelInputEl_" + type + "_" + head
            el.classList.add("checkExcelInputEl")
            el.classList.add("checkExcelInputEl_" + head)
            el.innerHTML = head
        $("#checkExcelInput" + type).appendChild(el)
    })
})

//파일 업로드 헤더 체크
let checkExcelHeader = (type, file) => {
    //첫 번째 줄을 체크
    let line = file[0]
    dataObj.step1[type].noHeader = 0
    headerList.forEach((head) => {
        //차후 생성되는 인풋은 무시
        if (headerNotInputList.indexOf(head) < 0) {
            if (line[head] !== undefined) {
                $("#checkExcelInputEl_" + type + "_" + head).classList.remove("no")
                $("#checkExcelInputEl_" + type + "_" + head).classList.add("yes")
            } else {
                dataObj.step1[type].noHeader += 1;
                $("#checkExcelInputEl_" + type + "_" + head).classList.remove("yes")
                $("#checkExcelInputEl_" + type + "_" + head).classList.add("no")
            }
        }
    })
}
let clearCheckExcelHeader = (type) => {
    let elArr = $("#checkExcelInput" + type).childNodes
    elArr.forEach((ele) => {
        ele.classList.remove("yes")
        ele.classList.remove("no")
    })
}

//파일 업로드
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
                    fileName:file.name,
                    name:file.name.replace(/\.[^/.]+$/, ""),
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
        //읽기 함수 병렬실행
        dataObj.step1.upload = []
        dataObj.step1.upload = await Promise.all(readArr)

        //업로드 된 파일 분석
        dataObj.step1.upload.forEach(file => {
            let firstLineHeader = checkFirstLineHeader(file)
            //비치희망 업로드
            if (firstLineHeader === false) {
                dataObj.step1.B.data = parseExcel(file)
                //열 정보 참조 수정
                dataObj.step1.B.data.forEach(line => {
                    referList.forEach(refer => {
                        if (line[refer[0]] !== undefined)
                            line[refer[1]] = line[refer[0]]
                    })
                })
                dataObj.step1.B.fileName = file.fileName
                dataObj.step1.B.name = file.name
                $("#uploadB").innerHTML = dataObj.step1.B.fileName + " (" + dataObj.step1.B.data.length + "건)"
                $("#uploadB").classList.add("uploaded")
                checkExcelHeader("B",dataObj.step1.B.data)//인풋 헤더 체크
            //동네서점 업로드
            } else {
                dataObj.step1.D.data = parseExcel(file)
                //열 정보 참조 수정
                dataObj.step1.D.data.forEach(line => {
                    referList.forEach(refer => {
                        if (line[refer[0]] !== undefined)
                            line[refer[1]] = line[refer[0]]
                    })
                })
                dataObj.step1.D.fileName = file.fileName
                dataObj.step1.D.name = file.name
                $("#uploadD").innerHTML = dataObj.step1.D.fileName + " (" + dataObj.step1.D.data.length + "건)"
                $("#uploadD").classList.add("uploaded")
                checkExcelHeader("D",dataObj.step1.D.data)//인풋 헤더 체크
            }
        })
        delete dataObj.step1.upload//업로드용 임시 배열 삭제
    }
}

//업로드 파일 삭제 버튼
$("#deleteB").onclick = () => {
    if (dataObj.step1.B.data !== undefined) {
        dataObj.step1.B = {}//데이터 초기화
        $("#uploadB").innerHTML = "(업로드 대기 중)"
        $("#uploadB").classList.remove("uploaded")
        clearCheckExcelHeader("B")//인풋 헤더 체크 초기화
    }
}
$("#deleteD").onclick = () => {
    if (dataObj.step1.D !== undefined) {
        dataObj.step1.D = {}//데이터 초기화
        $("#uploadD").innerHTML = "(업로드 대기 중)"
        $("#uploadD").classList.remove("uploaded")
        clearCheckExcelHeader("D")//인풋 헤더 체크 초기화
    }
}

//업로드 파일 서로 바꾸기 (업로드 에러 대비용)
$("#exchangeBD").onclick = () => {
    //바꾸기 의사 물어보기
    if (!confirm("\"비치희망신청\"과 \"동네서점바로대출\" 엑셀을 서로 바꾸겠습니까?")) return
    //바꿔치기
    let tempObj = JSON.parse(JSON.stringify(dataObj.step1.D))
    dataObj.step1.D = JSON.parse(JSON.stringify(dataObj.step1.B))
    dataObj.step1.B = JSON.parse(JSON.stringify(tempObj))
    
    //표시 변경
    if (Object.keys(dataObj.step1.B).length > 0) {
        $("#uploadB").innerHTML = dataObj.step1.B.fileName + " (" + dataObj.step1.B.data.length + "건)"
        $("#uploadB").classList.add("uploaded")
        checkExcelHeader("B",dataObj.step1.B.data)//인풋 헤더 체크
    } else {
        dataObj.step1.B = {}//데이터 초기화
        $("#uploadB").innerHTML = "(업로드 대기 중)"
        $("#uploadB").classList.remove("uploaded")
        clearCheckExcelHeader("B")//인풋 헤더 체크 초기화
    }
    if (Object.keys(dataObj.step1.D).length > 0) {
        $("#uploadD").innerHTML = dataObj.step1.D.fileName + " (" + dataObj.step1.D.data.length + "건)"
        $("#uploadD").classList.add("uploaded")
        checkExcelHeader("D",dataObj.step1.D.data)//인풋 헤더 체크
    } else {
        dataObj.step1.D = {}//데이터 초기화
        $("#uploadD").innerHTML = "(업로드 대기 중)"
        $("#uploadD").classList.remove("uploaded")
        clearCheckExcelHeader("D")//인풋 헤더 체크 초기화
    }
}

//다음 단계 버튼
$("#goNext1").onclick = () => {
    //파일 업로드 여부 체크
    if (dataObj.step1.B.data === undefined && dataObj.step1.D.data !== undefined) {
        if (!confirm("※ 비치희망신청 엑셀이 업로드되지 않았습니다.\n비치희망신청 건수를 0건으로 간주합니다.\n\n이대로 진행하겠습니까?")) {
          return
        }
    } else if (dataObj.step1.B.data !== undefined && dataObj.step1.D.data === undefined) {
        if (!confirm("※ 동네서점바로대출 엑셀이 업로드되지 않았습니다.\n동네서점바로대출 건수를 0건으로 간주합니다.\n\n이대로 진행하겠습니까?")) {
          return
        }
    } else if (dataObj.step1.B.data === undefined && dataObj.step1.D.data === undefined) {
        alert("※ 오류 : 업로드된 엑셀 파일이 없습니다.")
        return
    }
    //파일 헤더 체크
    if (dataObj.step1.B.noHeader !== undefined && dataObj.step1.B.noHeader > 0) {
        if (!confirm("※ 비치희망신청 엑셀에 없는 정보가 있습니다.\n(업로드 칸 상단 주황색 표시)\n\n그래도 진행하겠습니까?")) {
          return
        }
    }
    if (dataObj.step1.D.noHeader !== undefined && dataObj.step1.D.noHeader > 0) {
        if (!confirm("※ 동네서점바로대출 엑셀에 없는 정보가 있습니다.\n(업로드 칸 상단 주황색 표시)\n그래도 진행하겠습니까?")) {
          return
        }
    }

    //엑셀 업로드 막기, dropeZone 변경
    allowInput = false

    //다음 단계로 이동
    $("#page1").style.display = "none"
    $("#page2").style.display = "block"

    //파일 합치기 및 ISBN 분석 실시
    mergeAndCheckISBN(true)
}

//=============================
// 2. ISBN 중복 조회
//=============================
//A. 엑셀 합치기
let mergeAndCheckISBN = (goUp) => {
    //현재 및 후속 단계 정보 초기화
    dataObj.step2 = {merged:[],filtered:[],delete:[],deletable:[]}
    dataObj.step3 = {}
    dataObj.step4 = {}

    //단일 엑셀 정보 만들기
    let tempArr = []
      if (dataObj.step1.B.data !== undefined) tempArr.push(dataObj.step1.B.data)
      if (dataObj.step1.D.data !== undefined) tempArr.push(dataObj.step1.D.data)
    let num = 1
    tempArr.forEach((data,i) => {
        data.forEach((uploadR,j) => {
            let inputR = {}
            headerList.forEach(header => {
                //번호 : 순서대로 작성
                if (header === "번호") {
                    inputR[header] = num.toString()
                //일반 정보 : 옮겨적기
                } else if (uploadR[header] !== undefined) {
                    inputR[header] = uploadR[header]
                //구분 : 비치 or 동네 입력
                } else if (header === "구분") {
                    if (i === 0) {
                        inputR[header] = "비치"
                    } else {
                        inputR[header] = "동네"
                    }
                //없는 정보는 공란
                } else {
                    inputR[header] = ""
                }
            })
            //열 번호 상승
            num += 1
            //열 입력
            dataObj.step2.merged.push(inputR)
        })
    })

    //1차 필터링 (ISBN, 서명 중복 조회 전 - 불필요 상태 자료 제거)
    dataObj.step2.merged = dataObj.step2.merged.filter(function(row) {
        return ignoreStateList[0].indexOf(row["상태"]) < 0
    })

    //ISBN 동일 자료 찾기
    dataObj.step2.merged.forEach((i,num1) => {
        dataObj.step2.merged.forEach((j,num2) => {
            if (i["번호"] !== j["번호"] && i.ISBN.length > 0 && i.ISBN === j.ISBN) {
                if (dataObj.step2.filtered.indexOf(num1) < 0) {
                    dataObj.step2.filtered.push(num1)
                }
                if (dataObj.step2.filtered.indexOf(num2) < 0) {
                    dataObj.step2.filtered.push(num2)
                }
            }
        })
    })

    //출력용 표 생성
    let table = $("#dataTable2")
    //기존 데이터 비우기
    while (table.hasChildNodes()) {
        table.removeChild(table.lastChild);
    }
    //a. 헤더 추가
    let header = table.insertRow(-1)
    headerList.forEach(col => {
        let newCol = document.createElement("th")
        newCol.innerText = col
        //칸 아이디, 클래스 입력
        newCol.id = "dataTable2_header_" + col
        newCol.classList.add("dataTable_" + col)
        //헤더 명시
        newCol.classList.add("header")
        header.appendChild(newCol)
    })
    //b. 출력할 ISBN 중복 자료가 있을 경우 : 데이터 행 추가
    if (dataObj.step2.filtered.length > 0) {
        dataObj.step2.filtered.forEach((num,i) => {
            let newRow = table.insertRow(-1)
            //행 아이디 입력
            newRow.id = "dataTable2_" + dataObj.step2.merged[num]["번호"]
            //대응 번호 기록
            newRow.dataset.number = dataObj.step2.merged[num]["번호"]
            //해당 행이 신청 or 신청중 : 선택하여 삭제 예정 지정 가능
            if (dataObj.step2.merged[num]["상태"] === "신청" || dataObj.step2.merged[num]["상태"] === "신청중") {
                //삭제 가능 대상 기록
                dataObj.step2.deletable.push(num)
                //삭제 예정 기능
                newRow.onclick = () => {
                    let num = newRow.dataset.number
                    //삭제 예정 기록
                    if (dataObj.step2.delete.indexOf(num) < 0) {
                        dataObj.step2.delete.push(num)
                        $("#dataTable2_" + num.toString()).classList.add("willDelete")
                        $("#dataTable2_" + num.toString() + "_검토").innerHTML = "<span style='color:red;'>삭제</span>"
                    //삭제 예정 삭제
                    } else {
                        let index = dataObj.step2.delete.indexOf(num)
                        dataObj.step2.delete.splice(index, 1)
                        $("#dataTable2_" + num.toString()).classList.remove("willDelete")
                        $("#dataTable2_" + num.toString() + "_검토").innerHTML = ""
                    }
                    //삭제 예정 현황 갱신
                    $("#currentDelete2").value = "삭제 예정 : " + dataObj.step2.delete.length.toString() + " / " + dataObj.step2.deletable.length.toString()
                }
            }
            headerList.forEach((col) => {
                let newCol = newRow.insertCell()
                //칸 아이디, 클래스 입력
                newCol.id = "dataTable2_" + dataObj.step2.merged[num]["번호"] + "_" + col
                newCol.classList.add("dataTable_" + col)
                //내용 입력
                newCol.innerText = dataObj.step2.merged[num][col]
                    //번호는 별도로 매기기
                    if (col === "번호") newCol.innerText = (i+1).toString()
                //바디 명시
                newCol.classList.add("body")
                //ISBN이 달라질 때마다 상단 굵게
                if (i > 1 && dataObj.step2.merged[dataObj.step2.filtered[i]]["ISBN"] !== dataObj.step2.merged[dataObj.step2.filtered[i-1]]["ISBN"]) {
                    newCol.classList.add("seperatorBlack")
                }
                //신청 or 신청이 아니면 선택 불가
                if (dataObj.step2.merged[num]["상태"] !== "신청" && dataObj.step2.merged[num]["상태"] !== "신청중") {
                    newCol.classList.add("notDeletable")
                    //검토에 "처리 중" 표기
                    if (col === "검토") newCol.innerText = "진행"
                }
            })
        })
        //표 보여주기
        table.classList.remove("hidden")
        //삭제 예정 현황 표시
        $("#currentDelete2").value = "삭제 예정 : " + dataObj.step2.delete.length.toString() + " / " + dataObj.step2.deletable.length.toString()
        //중복 없음 문구 숨기기
        $("#noISBNDuplicate").style.display = "none"
    } else {
        //표 숨기기
        table.classList.add("hidden")
        //중복 없음 문구 표시
        $("#noISBNDuplicate").style.display = "block"
    }

    //페이지 최상단 스크롤 (희망 시)
    if (goUp === true) {
        $("#pageBody2").scrollTop = 0
    }
}

//삭제 예정 초기화 : 표 다시 생성
$("#restoreTable2").onclick = () => {
    //초기화 의사 물어보기
    if (!confirm("삭제 예정 선택을 초기화하겠습니까?")) return

    //초기화 실시
    mergeAndCheckISBN(false)
}

//이전 단계 버튼
$("#goPrevious2").onclick = () => {
    //엑셀 업로드 허용, dropeZone 변경
    allowInput = true

    $("#page2").style.display = "none"
    $("#page1").style.display = "block"
}
//다음 단계 버튼
$("#goNext2").onclick = () => {
    $("#page2").style.display = "none"
    $("#page3").style.display = "block"

    //서명 중복 조회 실시
    checkTitle(true)
}
//=============================
// 3. 서명 중복 조회
//=============================
let checkTitle = (goUp) => {
    //현재 및 후속 단계 정보 초기화
    dataObj.step3 = {merged:[],delete:[],deletable:[]}
    dataObj.step4 = {}

    //이전 단계 데이터 받아오기 (삭제 예정 제외)
    dataObj.step2.merged.forEach((row) => {
        if (dataObj.step2.delete.indexOf(row["번호"]) < 0) {
            dataObj.step3.merged.push(row)
        }
    })

    //서명으로 데이터 정렬
    dataObj.step3.merged.sort((a,b) => {
        let Key = [a["서명"],b["서명"]]

        if (Key[0] < Key[1]) {
            return -1
        } else if (Key[0] > Key[1]) {
            return 1
        } else {
            return 0
        }
    })

    //출력용 표 생성
    let table = $("#dataTable3")
    //기존 데이터 비우기
    while (table.hasChildNodes()) {
        table.removeChild(table.lastChild);
    }
    //a. 헤더 추가
    let header = table.insertRow(-1)
    headerList.forEach(col => {
        let newCol = document.createElement("th")
        newCol.innerText = col
        //칸 아이디, 클래스 입력
        newCol.id = "dataTable3_header_" + col
        newCol.classList.add("dataTable_" + col)
        //헤더 명시
        newCol.classList.add("header")
        header.appendChild(newCol)
    })
    //b. 데이터 행 추가
    dataObj.step3.merged.forEach((line,i) => {
        let newRow = table.insertRow(-1)
        //행 아이디 입력
        newRow.id = "dataTable3_" + line["번호"]
        //대응 번호 기록
        newRow.dataset.number = line["번호"]
        //해당 행이 신청 or 신청 : 선택하여 삭제 예정 지정 가능
        if (line["상태"] === "신청" || line["상태"] === "신청중") {
            //삭제 가능 대상 기록
            dataObj.step3.deletable.push(line["번호"])
            //삭제 예정 기능
            newRow.onclick = () => {
                let num = newRow.dataset.number
                //삭제 예정 기록
                if (dataObj.step3.delete.indexOf(num) < 0) {
                    dataObj.step3.delete.push(num)
                    $("#dataTable3_" + num.toString()).classList.add("willDelete")
                    $("#dataTable3_" + num.toString() + "_검토").innerHTML = "<span style='color:red;'>삭제</span>"
                //삭제 예정 삭제
                } else {
                    let index = dataObj.step3.delete.indexOf(num)
                    dataObj.step3.delete.splice(index, 1)
                    $("#dataTable3_" + num.toString()).classList.remove("willDelete")
                    $("#dataTable3_" + num.toString() + "_검토").innerHTML = ""
                }
                //삭제 예정 현황 갱신
                $("#currentDelete3").value = "삭제 예정 : " + dataObj.step3.delete.length.toString() + " / " + dataObj.step3.deletable.length.toString()
            }
        }
        headerList.forEach((col,j) => {
            let newCol = newRow.insertCell()
            //칸 아이디, 클래스 입력
            newCol.id = "dataTable3_" + line["번호"] + "_" + col
            newCol.classList.add("dataTable_" + col)
            //내용 입력
                //번호 : 순서대로 매기기
                if (col === "번호") {
                    newCol.innerText = (i+1).toString()
                //서명 : 첫 글자 ~ 위/아래 같은 글자까지 크게 강조
                } else if (col === "서명") {
                    let titleStr = (' ' + line[col]).slice(1)
                    let bigEnd = 0
                    for (let k = 0;k<titleStr.length;k++) {
                        if ((dataObj.step3.merged[i-1] !== undefined && dataObj.step3.merged[i]["서명"][k] === dataObj.step3.merged[i-1]["서명"][k]) ||
                            (dataObj.step3.merged[i+1] !== undefined && dataObj.step3.merged[i]["서명"][k] === dataObj.step3.merged[i+1]["서명"][k])) {
                                bigEnd += 1
                        } else break
                    }
                    if (bigEnd === 0) bigEnd = 1
                    titleStr = titleStr.slice(0,bigEnd) + "</span>" + titleStr.slice(bigEnd)
                    newCol.innerHTML = "<span class='big'>" + titleStr
                //나머지 : 그대로 내용 입력
                } else {
                    newCol.innerText = line[col]
                }
            //바디 명시
            newCol.classList.add("body")
            //신청 or 신청이 아니면 선택 불가
            if (line["상태"] !== "신청" && line["상태"] !== "신청중") {
                newCol.classList.add("notDeletable")
                //검토에 "처리 중" 표기
                if (col === "검토") newCol.innerText = "진행"
            }
        })
    })
    //표 보여주기
    table.classList.remove("hidden")
    //삭제 예정 현황 표시
    $("#currentDelete3").value = "삭제 예정 : " + dataObj.step3.delete.length.toString() + " / " + dataObj.step3.deletable.length.toString()

    //페이지 최상단 스크롤 (희망 시)
    if (goUp === true) {
        $("#pageBody3").scrollTop = 0
    }
}

//삭제 예정 초기화 : 표 다시 생성
$("#restoreTable3").onclick = () => {
    //초기화 의사 물어보기
    if (!confirm("삭제 예정 선택을 초기화하겠습니까?")) return

    //초기화 실시
    checkTitle(false)
}

//이전 단계 버튼
$("#goPrevious3").onclick = () => {
    $("#page3").style.display = "none"
    $("#page2").style.display = "block"
}
//다음 단계 버튼
$("#goNext3").onclick = () => {
    $("#page3").style.display = "none"
    $("#page6").style.display = "block"
    $("#page5").style.display = "block"
    $("#page4").style.display = "block"

    //엑셀 출력 전 2차 필터링
    secondFilter()
}
//=============================
// 4. 반입 & 검토 시작
//=============================
//금일 날짜 계산 (파일 이름)
let today = new Date()
let todayText = " (" + today.getFullYear().toString() + "." + ("0" + (today.getMonth() + 1).toString()).slice(-2) + "." + ("0" + today.getDate().toString()).slice(-2) + ")"

//엑셀 출력 전 2차 필터링
let secondFilter = () => {
    //현재 단계 정보 초기화
    dataObj.step4 = []

    //이전 단계 데이터 받아오기 (삭제 예정 제외)
    dataObj.step3.merged.forEach((row) => {
        if (dataObj.step3.delete.indexOf(row["번호"]) < 0) {
            dataObj.step4.push(row)
        }
    })

    //2차 필터링 (ISBN, 서명 중복 조회 후 - 신청 단계 자료만 남기기)
    dataObj.step4 = dataObj.step4.filter(function(row) {
        return ignoreStateList[1].indexOf(row["상태"]) < 0
    })

    //번호로 데이터 정렬
    dataObj.step4.sort((a,b) => {
        let Key = [parseInt(a["번호"]),parseInt(b["번호"])]

        if (Key[0] < Key[1]) {
            return -1
        } else if (Key[0] > Key[1]) {
            return 1
        } else {
            return 0
        }
    })

    //반입용 엑셀 생성
    excelForImport()
}

//반입용 엑셀 출력 준비
let excelForImport = () => {
    //출력용 표 생성
    let table = $("#dataTable4")
    //기존 데이터 비우기
    while (table.hasChildNodes()) {
        table.removeChild(table.lastChild);
    }
    //a. 헤더 추가
    let header = table.insertRow(-1)
    importHeaderList.forEach(col => {
        let newCol = document.createElement("th")
        newCol.innerText = col
        //칸 클래스 입력
        newCol.classList.add("dataTable_" + col.replaceAll(" ","_"))
        //헤더 명시
        newCol.classList.add("header")
        header.appendChild(newCol)
    })
    dataObj.step4.forEach((line,i) => {
        let newRow = table.insertRow(-1)
        //대응 번호 기록
        importHeaderList.forEach((col) => {
            let newCol = newRow.insertCell()
            //칸 클래스 입력
            newCol.classList.add("dataTable_" + col.replaceAll(" ","_"))
            //바디 명시
            newCol.classList.add("body")
            //(대응 내용이 있으면) 입력
            if (col in referImportList) {
                newCol.innerText = line[referImportList[col]]
            }
            //제목은 표제만 추출하기
            if (col === "서명") {
                newCol.innerText = line[col].replaceAll("-","(").replaceAll(":","(").split("(")[0]
            }
            //면장수 입력하기
            if (col === "면장수") {
                newCol.innerText = "1책"
            }
        })
    })
    //표 보여주기
    table.classList.remove("hidden")

    //엑셀 출력 버튼 세팅
    $("#writeExcelImport").onclick = () => {
        writeExcel($("#dataTable4"),"반입용 엑셀" + todayText + ".xlsx",
            {colWidth://열 너비
                [
                    40,//서명(필수)
                    20,//저자(필수)
                    20,//발행자(필수)
                    10,//발행년(필수)
                    20,//낱권ISBN(입력)
                    5,//낱권ISBN부가기호
                    10,//가격(필수)
                    5,//세트ISBN
                    5,//세트ISBN부가기호
                    5,//KDC 분류기호
                    5,//DDC 분류기호
                    5,//기타 분류기호
                    5,//청구기호_분류
                    5,//청구기호_도서기호
                    5,//청구기호_권책기호
                    5,//부서명
                    5,//대등서명
                    5,//공저자
                    5,//판사항
                    5,//편/권차
                    5,//편제
                    5,//발행지
                    8,//면장수(필수)
                    5,//물리적특성
                    5,//크기
                    5,//딸림자료
                    5,//총서표제
                    5,//총서편차
                    5,//수상주기
                    5,//등록번호
                    5,//복본기호
                    5,//별치기호
                    5//URI
                ],
            hideCol://열 숨기기
                [5,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,23,24,25,26,27,28,29,30,31,32]
            }
        )
    }

    //검토용 엑셀 생성
    excelForDetermine()
}

//검토용 엑셀 출력 준비
let excelForDetermine = () => {
    //출력용 표 생성
    let table = $("#dataTable5")
    //기존 데이터 비우기
    while (table.hasChildNodes()) {
        table.removeChild(table.lastChild);
    }
    //a. 헤더 추가
    let header = table.insertRow(-1)
    headerList.forEach(col => {
        //상태는 헤더에서 제외 (어차피 전부 신청 상태)
        if (col !== "상태") {
            let newCol = document.createElement("th")
            newCol.innerText = col
            //칸 클래스 입력
            newCol.classList.add("dataTable_" + col.replaceAll(" ","_"))
            //헤더 명시
            newCol.classList.add("header")
            header.appendChild(newCol)
        }
    })
    dataObj.step4.forEach((line,i) => {
        let newRow = table.insertRow(-1)
        //행 아이디 입력
        newRow.id = "dataTable5_" + line["번호"]
        //대응 번호 기록
        newRow.dataset.number = line["번호"]
        //행 입력
        headerList.forEach((col) => {
            //상태는 표에서 제외 (어차피 전부 신청 상태)
            if (col !== "상태") {
                let newCol = newRow.insertCell()
                //칸 클래스 입력
                newCol.classList.add("dataTable_" + col.replaceAll(" ","_"))
                //바디 명시
                newCol.classList.add("body")
                //내용 입력
                    //번호 : 행 순번 입력 & "숫자" 타입
                    if (col === "번호") {
                        newCol.innerText = (i+1).toString()
                        newCol.dataset.type = "number"
                    //가격 : 천단위 구분 추가 & "숫자" 타입
                    } else if (col === "가격") {
                        newCol.innerText = line[col]
                        newCol.dataset.type = "currency"
                    //나머지 : 내용 그대로 입력
                    } else {
                        newCol.innerText = line[col]
                    }
                //비치희망, 동네서점 빨강 구분선
                if (i > 0 && dataObj.step4[i-1]["구분"] === "비치" && dataObj.step4[i]["구분"] === "동네") {
                    newCol.classList.add("seperatorRed")
                }
            }
        })
    })
    //표 보여주기
    table.classList.remove("hidden")

    //알라딘 검색 제외용 표 생성 준비(검토 엑셀 표 복사해오기)
    let createSearchRejectTable = () => {
        $("#dataTable6").innerHTML = ""
        $("#dataTable6").innerHTML = table.innerHTML

        //검색 제외 준비
        dataObj.step4.reject = []
        //검색 제외 기능
        for (let i = 0, row; row = $("#dataTable6").rows[i]; i++) {
            row.id = row.id.replace("dataTable5","dataTable6")
            let num = row.dataset.number
            row.onclick = () => {
                //검색 제외 기록
                if (dataObj.step4.reject.indexOf(num) < 0) {
                    dataObj.step4.reject.push(num)
                    $("#dataTable6_" + num.toString()).classList.add("reject")
                    $("#dataTable6_" + num.toString() + " .dataTable_검토").innerHTML = "<span style='color:red;'>제외</span>"
                //검색 제외 삭제
                } else {
                    let index = dataObj.step4.reject.indexOf(num)
                    dataObj.step4.reject.splice(index, 1)
                    $("#dataTable6_" + num.toString()).classList.remove("reject")
                    $("#dataTable6_" + num.toString() + " .dataTable_검토").innerHTML = ""
                }
            }
        }
    }

    //알라딘 검색 제외용 표 생성
    createSearchRejectTable()

    //검색 제외 초기화 : 표 다시 생성
    $("#restoreTable6").onclick = () => {
        //초기화 의사 물어보기
        if (!confirm("검색 제외 선택을 초기화하겠습니까?")) return

        //초기화 실시
        createSearchRejectTable()
    }
    
    //엑셀 출력 버튼 세팅
    $("#writeExcelConsider").onclick = () => {
        writeExcel($("#dataTable5"),"검토용 엑셀" + todayText + ".xlsx",
            {colWidth://열 너비
                [
                    6,//번호
                    6,//구분
                    17,//ISBN
                    6,//검토
                    30,//서명
                    18,//저자
                    18,//발행자
                    8,//발행년
                    8,//가격
                    18,//대출자번호
                    10,//신청자
                    22//신청일
                ]
            }
        )
    }

    //알라딘 검색 창 열기
    let openAladin = async (book) => {
        //검색 제외 대상일 경우 패스
        if (dataObj.step4.reject.indexOf(book["번호"]) >= 0) return
        //0.01초 단위로 창 열기 (순차적으로 열어야 창이 생성됨)
        let link = "https://www.aladin.co.kr/search/wsearchresult.aspx?SearchTarget=All&SearchWord=" + book.ISBN
        window.open(link, "_blank")
        await sleep(10)
    }

    //알라딘 일괄검색 버튼
    $("#searchAladinUp").onclick = async () => {
        for(let i = 0;i<dataObj.step4.length;i++) {
            openAladin(dataObj.step4[i])
        }
    }
    
    $("#searchAladinDown").onclick = async () => {
        for(let i = dataObj.step4.length - 1;i>= 0;i--) {
            openAladin(dataObj.step4[i])
        }
    }
}

//이전 단계 버튼
$("#goPrevious6").onclick = () => {
    $("#page4").style.display = "none"
    $("#page5").style.display = "none"
    $("#page6").style.display = "none"
    $("#page3").style.display = "block"
}