const xlsx = require('xlsx'); //엑셀 모듈 가져옴

const excelFile = xlsx.readFile('src/common/투비펫.xlsx');
const sheetName = excelFile.SheetNames[0];
const firstSheet = excelFile.Sheets[sheetName];
const tempData = xlsx.utils.sheet_to_json(firstSheet, {defval: ''});
const list = xlsx.utils.book_new();

const jsonParser = (array) => {
    //CJ송장양식 //commit테스트
    const sellList = [
        [
            '주문번호',
            '상품명',
            '옵션',
            '수량',
            '배송료',
            '배송방법',
            '주문자', 
            '주문자전화',
            '주문자핸드폰',
            '수령자',
            '전화', 
            '핸드폰',
            '우편번호',
            '주소',
            '배송메세지',
          ],
    ];

    const jsonData = array.map(v => {
        const newRecord = {};
        for (const key in v) {
            const newString = v[key]+"";
            const isSize = newString.indexOf("사이즈선택")
            newRecord[key.replace(/\s/g, "")] = v[key]; //공백 제거
            if(isSize != -1) newRecord[key] = v[key].replace("사이즈선택", v.__EMPTY_7) //상품명의 사이즈선택을 옵션에있는 사이즈로 대체
        }
        return newRecord
    })

    //판매명부 작성
    jsonData.map((v, i) => {
        if(i !== 0) {
        const newList = [];

        newList.push('__AUTO__'); //주문번호
        newList.push('투) '+v.__EMPTY_6); //상품명
        newList.push(''); //옵션
        newList.push(v.__EMPTY_8); //수량
        newList.push(2500); //배송료
        newList.push('선결제'); //배송방법
        newList.push(v.__EMPTY_3); //주문자
        newList.push(v.__EMPTY_4 || v.__EMPTY_5); //주문자전화
        newList.push(v.__EMPTY_5 || v.__EMPTY_4); //주문자핸드폰
        newList.push(v.__EMPTY_3); //수령자
        newList.push(v.__EMPTY_4 || v.__EMPTY_5); //전화
        newList.push(v.__EMPTY_5 || v.__EMPTY_4); //핸드폰
        newList.push(v.__EMPTY_20); //우편번호
        newList.push(v.__EMPTY_21); //주소
        newList.push(v.__EMPTY_22); //배송메세지

        sellList.push(newList)
        }
    })
    return sellList
}

//엑셀에 넣을 데이터 생성
const sellData = jsonParser(tempData);

//엑셀 형식에 맞게 시트 데이터 생성
const orders = xlsx.utils.aoa_to_sheet(sellData);


// 셀 크기 지정
orders['!cols'] = [
    { wpx: 90 }, // A열
  { wpx: 300 }, // B열
  { wpx: 30 }, // C열
  { wpx: 30 }, // D열
  { wpx: 50 }, // E열
  { wpx: 50 }, // F열
  { wpx: 200 }, // G열
  { wpx: 100 }, // H열
  { wpx: 100 }, // I열
  { wpx: 200 }, // J열
  { wpx: 100 }, // K열
  { wpx: 100 }, // L열   
  { wpx: 100 }, // M열
  { wpx: 200 }, // N열
]

//시트 생성
xlsx.utils.book_append_sheet(list, orders, '투비펫');

//엑셀 파일 생성
xlsx.writeFile(list, '투비펫송장.xlsx');