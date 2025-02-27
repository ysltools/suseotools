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

//허용 확장자 : txt
let allowedExtensionsReg = /(\.txt)$/i
let allowedExtensionsArr = ["txt"]

//데이터 관리 객체
let upload//입력 파일 관리
let inputArr = []//입력값 관리

//엑셀 업로드 제어
let allowInput = true


//=============================
// 1. 파일 업로드
//=============================
//A. 드래그 방식
let dropZone = $("#dropZone")

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
        dropZone.innerHTML = "여기에 .TXT 파일을 올리세요"
        dropZone.classList.add("show1")
    } else {
        dropZone.innerHTML = "지금은 .TXT 파일을 올릴 수 없습니다"
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
                resolve(e.target.result)
            }
            reader.readAsText(file)
        })
    }
    //첫 번째 파일만 입력하기
    let file = uploadArr[0]
    //허용하지 않는 확장자일 경우, 입력 취소
    if(!allowedExtensionsReg.exec(file.name)){
        alert("올바른 [.txt] 파일이 입력되지 않았습니다.")
        return
    } else {
        upload = await readFile(file)
        $("#inputArea").value = upload
        //입력값 확인
        getInput()
    }
}

//=============================
// 2. 입력값 저장 및 초기화
//=============================
//함수 : 입력값이 변경될 때마다 저장
let getInput = () => {
    let text = $("#inputArea").value
    inputArr = text.split("\n").map((item) => {return item.trim()}).filter(i => i)
    /*
    //입력값 저장 검수용
    if (inputArr.length > 0) {
        console.log("<현재 입력값>\n" + inputArr.join("\n"))
    } else {
        console.log("<현재 입력값>\n(없음)")
    }
    */
    $("#bookNum").innerHTML = inputArr.length.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",")
}

//입력값이 변경될 때마다 확인
$("#inputArea").oninput = getInput

//입력값 초기화 버튼
$("#clearInput").onclick = () => {
    let confirm = window.confirm("입력값을 초기화하겠습니까?")
    if (confirm) {
        $("#inputArea").value = ""
        inputArr = []
    }
    //입력값 확인
    getInput()
}

//=============================
// 3. 알라딘 검색
//=============================
//함수 : 알라딘 검색 창 열기
let openAladin = async (keyword) => {
    //0.02초 단위로 창 열기 (순차적으로 열어야 창이 생성됨)
    let link = "https://www.aladin.co.kr/search/wsearchresult.aspx?SearchTarget=All&SearchWord=" + encodeURIComponent(keyword)
    window.open(link, "_blank")
    await sleep(20)
}

//알라딘 일괄검색 버튼
$("#searchAladinUp").onclick = async () => {
    if (inputArr.length > 0) {
        for(let i = 0;i<inputArr.length;i++) {
            openAladin(inputArr[i])
        }
    }
}

$("#searchAladinDown").onclick = async () => {
    if (inputArr.length > 0) {
        for(let i = inputArr.length - 1;i>= 0;i--) {
            openAladin(inputArr[i])
        }
    }
}
