import requests
from Crypto.PublicKey import RSA
from Crypto.Cipher import PKCS1_v1_5
import base64
from bs4 import BeautifulSoup
import json
import export_curriculum_excel


def get_public_key(session, public_key_url):
    # 获取公钥
    public_key_response = session.get(public_key_url)

    if public_key_response.status_code == 200:
        try:
            public_key_data = public_key_response.json()
            modulus = public_key_data['modulus']
            exponent = public_key_data['exponent']
            return modulus, exponent
        except ValueError:
            return None, None
    else:
        return None, None


# 加密密码
def encrypt_password(password, modulus_b64, exponent_b64):
    modulus = int.from_bytes(base64.b64decode(modulus_b64), 'big')
    exponent = int.from_bytes(base64.b64decode(exponent_b64), 'big')
    public_key = RSA.construct((modulus, exponent))
    cipher = PKCS1_v1_5.new(public_key)
    encrypted_password = cipher.encrypt(password.encode())
    return base64.b64encode(encrypted_password).decode()


def load_config():
    with open('user_config.json', 'r', encoding='utf-8') as f:
        return json.load(f)


def get_kb():
    config = load_config()
    vpn_credentials = config['vpn_credentials']
    system_credentials = config['system_credentials']
    curriculum_param = config['curriculum_param']
    vpn_login_url = config['vpn_login_url']
    public_key_url = config['public_key_url']
    system_login_url = config['system_login_url']
    curriculum_url = config['curriculum_url']
    curriculum_html_url = config['curriculum_html_url']

    session = requests.Session()

    # VPN 登录
    login_page = session.get(vpn_login_url)
    soup = BeautifulSoup(login_page.text, 'html.parser')
    hidden_inputs = soup.find_all('input', type='hidden')
    form_data = {inp['name']: inp['value'] for inp in hidden_inputs if 'name' in inp.attrs}
    form_data.update(vpn_credentials)
    vpn_response = session.post(vpn_login_url, data=form_data)
    if "登录成功" not in vpn_response.text and vpn_response.url == vpn_login_url:
        raise Exception("VPN 登录失败")

    # 获取系统公钥
    modulus, exponent = get_public_key(session, public_key_url)
    if modulus is None or exponent is None:
        raise Exception("公钥获取失败")

    # 加密系统密码
    system_password = system_credentials['mm']
    encrypted_password = encrypt_password(system_password, modulus, exponent)
    system_credentials['mm'] = encrypted_password

    # 获取系统登录页面，找 CSRF
    system_login_page = session.get(system_login_url)
    soup = BeautifulSoup(system_login_page.text, 'html.parser')
    csrf_token_input = soup.find('input', {'name': 'csrftoken'})
    if csrf_token_input:
        csrftoken = csrf_token_input['value']
        system_credentials['csrftoken'] = csrftoken

    # 提交系统登录
    system_response = session.post(system_login_url, data=system_credentials)
    if system_response.status_code != 200:
        raise Exception("系统登录失败")

    # 获取课表 HTML 页面
    html_response = session.get(curriculum_html_url)
    if html_response.status_code != 200:
        raise Exception("获取课表 HTML 失败")

    soup = BeautifulSoup(html_response.text, 'html.parser')
    xnm = soup.find('input', {'id': 'xnm_hide'})['value']
    xqm = soup.find('input', {'id': 'xqm_hide'})['value']
    dqzc = soup.find('input', {'id': 'dqzc_hide'})['value']
    curriculum_param['xnm'] = xnm
    curriculum_param['xqm'] = xqm
    curriculum_param['zs'] = dqzc

    # 获取课表 JSON
    curriculum_response = session.post(curriculum_url, data=curriculum_param)
    if curriculum_response.status_code != 200:
        raise Exception("获取课表 JSON 失败")

    json_data = curriculum_response.json()
    if 'kbList' not in json_data:
        raise Exception("课表 JSON 中无 kbList")

    kb_list = json_data['kbList']

    # 过滤字段
    filtered_kb_list = [
        {key: item.get(key, "") for key in ['jxbmc', 'cdmc', 'jcor', 'xqj']}
        for item in kb_list
    ]

    return filtered_kb_list,dqzc


if __name__ == "__main__":
    kb,zs= get_kb()
    export_curriculum_excel.export(kb,f"第{zs}周课程表.xlsx")
