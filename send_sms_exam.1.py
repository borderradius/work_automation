'''
sms 보내기
'''
# cafe 24 php로 포팅해서 사용할 예정
# 아이디 : id
# secure 인증키 : ***************
# mode : '1' (base64 인코딩 모드)
# euckr 인코딩이 용량이 더 작아서 쓰는게 나음.
# avi, mpeg 등으로 인코딩할때 용량이 다른것처럼 euckr, utf8도 마찬가지.
# 인코딩시 짤린 문자열은 디코딩할때 깨짐. UnicodeDecodeError 를 출력.

from base64 import b64encode
import requests




'''
잔여건수조회
'''
# 인증정보
# url = "http://sslsms.cafe24.com/sms_remain.php"
# user_id = b64encode("sms id".encode("euckr"))
# secure = b64encode("********************".encode("euckr"))
# mode = b64encode("1".encode("euckr"))
# sender = '026451135'

# response = requests.post(url, data={
#     'user_id': user_id,
#     'secure': secure,
#     'mode': mode,
# })
# print(response.text) # 잔여 sms갯수


'''
sms 전송
'''
def send_sms(user_id, secure, sender, receivers, message):
    params = {
        'user_id': user_id,
        'secure': secure,
        'mode': '1',
        'sphone1': sender[:2],
        'sphone2': sender[2:5],
        'sphone3': sender[5:9],
        'rphone': ','.join(receivers),
        'msg': message,
    }

    data = {}
    for key, value in params.items():
        if isinstance(value, str):
            value = value.encode('euckr')
            if key == 'msg':
                value = value[:90].decode('euckr', 'ignore').encode('euckr')
        data[key] = value

    response = requests.post('https://sslsms.cafe24.com/sms_sender.php',data=data)
    return response.text


send_sms('sms id', 'secure key', '026451135',
         ['01066452135'], '전송될 문자 내용')
