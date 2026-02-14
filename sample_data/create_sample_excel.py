"""
이지어드민 발주서 형식의 샘플 엑셀 파일 생성 스크립트
"""
import pandas as pd
from datetime import datetime, timedelta
import random

# 샘플 데이터 생성
def create_sample_data():
    # 업체 리스트
    vendors = ['텀블러마켓', '에코용품사', '라이프스타일', '홈데코', '주방용품']

    # 상품 리스트
    products = {
        '텀블러마켓': [('스테인리스 텀블러', '블랙'), ('보온병', '화이트'), ('유리컵', '투명')],
        '에코용품사': [('대나무 칫솔', '내추럴'), ('친환경 빨대', '그린'), ('면 에코백', '베이지')],
        '라이프스타일': [('노트', '라인'), ('볼펜 세트', '블랙'), ('메모지', '핑크')],
        '홈데코': [('캔들', '라벤더향'), ('디퓨저', '시트러스'), ('액자', '우드')],
        '주방용품': [('실리콘 주걱', '레드'), ('냄비받침', '블루'), ('키친타올', '화이트')]
    }

    # 랜덤 이름 생성
    last_names = ['김', '이', '박', '최', '정', '강', '조', '윤', '장', '임']
    first_names = ['민준', '서연', '지우', '하은', '도윤', '예은', '시우', '서준', '지안', '수빈']

    # 랜덤 주소 생성
    cities = ['서울시 강남구', '서울시 서초구', '경기도 성남시', '경기도 용인시', '인천시 남동구',
              '부산시 해운대구', '대구시 수성구', '광주시 서구', '대전시 유성구', '울산시 남구']

    # 100개 주문 생성
    orders = []
    base_date = datetime(2024, 1, 28)

    for i in range(100):
        vendor = random.choice(vendors)
        product, option = random.choice(products[vendor])

        order = {
            '주문일자': (base_date + timedelta(days=random.randint(0, 2))).strftime('%Y-%m-%d'),
            '주문번호': f'ORD{base_date.strftime("%Y%m%d")}{i+1:04d}',
            '공급처': vendor,
            '수취인명': random.choice(last_names) + random.choice(first_names),
            '연락처': f'010-{random.randint(1000,9999)}-{random.randint(1000,9999)}',
            '주소': f'{random.choice(cities)} {random.choice(["가","나","다","라","마"])}동 {random.randint(1,999)}',
            '상품명': product,
            '옵션': option,
            '수량': random.randint(1, 3),
            '택배사': '',  # 업체가 입력할 칸
            '송장번호': ''  # 업체가 입력할 칸
        }
        orders.append(order)

    return pd.DataFrame(orders)

if __name__ == '__main__':
    # 샘플 데이터 생성
    df = create_sample_data()

    # 엑셀 파일로 저장
    output_file = '../input/이지어드민_발주서_20240128.xlsx'
    df.to_excel(output_file, index=False, engine='openpyxl')

    print(f'✅ 샘플 엑셀 파일 생성 완료: {output_file}')
    print(f'📊 총 주문 건수: {len(df)}')
    print(f'📦 업체별 주문 분포:')
    print(df['공급처'].value_counts())
