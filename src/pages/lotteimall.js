const xlsx = require('xlsx'); //엑셀 모듈 가져옴

const excelFile = xlsx.readFile('src/common/롯데홈쇼핑.xlsx');
const sheetName = excelFile.SheetNames[0];
const firstSheet = excelFile.Sheets[sheetName];
const tempData = xlsx.utils.sheet_to_json(firstSheet, {defval: ''});
const list = xlsx.utils.book_new();

const jsonParser = (array) => {
    //CJ송장양식
    const sellList = [
        [
            '주문일',
            '마스터상품코드',
            '상품코드',
            '주문번호',
            '상품명',
            '옵션',
            '수량',
            '판매가',
            '공급가',
            '원가',
            '추가구매옵션',
            '배송료',
            '배송방법',
            '주문자', 
            '주문자전화',
            '주문자핸드폰',
            '주문자이메일',
            '수령자',
            '전화', 
            '핸드폰',
            '수령자영문이름',
            '수령자주민등록번호(통관용)',
            '우편번호',
            '주소',
            '배송메세지',
            '배송사명',
            '송장번호',
            '사은품',
            '사용자임의분류'
          ],
    ];
    
    //키값에 공백이 있어서 공백 제거 및 특수문자 제거
    const jsonData = array.map(v => {
        const newRecord = {};
        for (const key in v) {
            const newString = v[key]+"";
            newRecord[key.replace(/\s/g, "")] = v[key]; //공백 제거 

            newRecord[key.replace(/[\{\}\[\]\/?.,;:|\)*~`!^\-+<>@\#$%&\\\=\(\'\"]/gi, "")] = v[key]; //특수문자 제거
        }
        return newRecord
    })
    
    //판매명부 작성
    jsonData.map(v => {
        const newList = [];

        newList.push('') //주문일
        newList.push('') //마스터상품코드
        newList.push('') //상품코드
        newList.push('__AUTO__'); //주문번호
        newList.push('롯)'+v.상품명); //상품명
        newList.push(''); //옵션
        newList.push(v.주문수량); //수량
        newList.push('') //판매가
        newList.push('') //공급가
        newList.push('') //원가
        newList.push('') //추가구매옵션
        newList.push(''); //배송료
        newList.push('선결제'); //배송방법
        newList.push(v.수취인); //주문자
        newList.push(v.수취인전화번호1); //주문자전화
        newList.push(v.수취인전화번호2); //주문자핸드폰
        newList.push('') //주문자이메일
        newList.push(v.수취인); //수령자
        newList.push(v.수취인전화번호1); //전화
        newList.push(v.수취인전화번호2); //핸드폰
        newList.push('') //수령자영문이름
        newList.push('') //수령자주민등록번호(통관용)
        newList.push(v.수취인우편번호); //우편번호
        newList.push(v.수취인주소); //주소
        newList.push(v.배송메시지); //배송메세지
        newList.push('') //배송사명
        newList.push('') //송장번호
        newList.push('') //사은품
        newList.push('') //사용자임의분류

        sellList.push(newList)
    })
    console.log(sellList)
    return sellList
}

//엑셀에 넣을 데이터 생성
const sellData = jsonParser(tempData);

//엑셀 형식에 맞게 시트 데이터 생성
const orders = xlsx.utils.aoa_to_sheet(sellData);


// 셀 크기 지정
orders['!cols'] = [
    { wpx: 0 }, // 주문일
    { wpx: 0 }, // 마스터상품코드
    { wpx: 0 }, // 상품코드
    { wpx: 90 }, // 주문번호
    { wpx: 300 }, // 상품명
    { wpx: 90 }, // 옵션
    { wpx: 30 }, // 수량
    { wpx: 0 }, // 판매가
    { wpx: 0 }, // 공급가
    { wpx: 0 }, // 원가
    { wpx: 0 }, // 추가구매옵션
    { wpx: 0 }, // 배송료
    { wpx: 50 }, // 배송방법
    { wpx: 50 }, // 주문자
    { wpx: 200 }, // 주문자전화
    { wpx: 100 }, // 주문자핸드폰
    { wpx: 0 }, // 주문자이메일
    { wpx: 100 }, // 수령자
    { wpx: 200 }, // 전화
    { wpx: 100 }, // 핸드폰
    { wpx: 0 }, // 수령자영문이름
    { wpx: 0 }, // 수령자주민등록번호(통관용)
    { wpx: 100 }, // 우편번호
    { wpx: 100 }, // 주소
    { wpx: 200 }, // 배송메세지
    { wpx: 0 }, // 배송사명
    { wpx: 0 }, // 송장번호
    { wpx: 0 }, // 사은품
    { wpx: 0 }, // 사용자임의분류
]

//시트 생성
xlsx.utils.book_append_sheet(list, orders, '롯데홈쇼핑송장');

//엑셀 파일 생성
xlsx.writeFile(list, '롯데홈쇼핑송장.xlsx');