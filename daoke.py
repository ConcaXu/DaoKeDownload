# 使用之前需要pip install PyMuPDF
# pip install requests
import requests
import os
import glob
import fitz
import html
from urllib import parse
import json
import time
import random
from requests.adapters import HTTPAdapter
from urllib3 import Retry

# 创建带重试机制的 Session（全局配置）
session = requests.Session()
retries = Retry(
    total=3,  # 最大重试次数
    backoff_factor=1,  # 重试等待时间：1, 2, 4秒
    status_forcelist=[500, 502, 503, 504]
)
session.mount('https://', HTTPAdapter(max_retries=retries))

# 统一请求头配置（模拟移动端浏览器）
COMMON_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Linux; Android 10; Pixel 3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.82 Mobile Safari/537.36',
    'Referer': 'https://m.doc88.com/',
    'Accept': 'image/webp,image/apng,image/*,*/*;q=0.8',
    'Accept-Encoding': 'gzip, deflate, br',
    'Accept-Language': 'zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7'
}


# 解码函数（保持原样）
# ... [此处保持原有 decode_base64 函数代码不变] ...

def download_gif(url: str, folder: str, filename: str, max_retry=3):
    """增强版下载函数，包含重试和随机延迟"""
    for attempt in range(max_retry):
        try:
            # 随机延迟（1~5秒）
            time.sleep(random.uniform(1, 5))

            response = session.get(
                url,
                headers=COMMON_HEADERS,
                timeout=15,  # 超时时间
                verify=False  # 临时关闭SSL验证（解决证书问题）
            )
            response.raise_for_status()  # 检查HTTP状态码

            # 写入文件
            with open(os.path.join(folder, filename), 'wb') as f:
                f.write(response.content)
            return True
        except Exception as e:
            print(f"第 {attempt + 1} 次下载失败: {e}")
            if attempt == max_retry - 1:
                return False


def pic2pdf(file_name):
    """合并PDF函数（优化内存管理）"""
    doc = fitz.open()
    try:
        img_list = sorted(glob.glob(os.path.join(file_name, '*.gif')))
        if not img_list:
            raise ValueError("未找到GIF文件")

        for img_path in img_list:
            imgdoc = fitz.open(img_path)
            pdfbytes = imgdoc.convert_to_pdf()
            imgpdf = fitz.open("pdf", pdfbytes)
            doc.insert_pdf(imgpdf)
            imgdoc.close()
            imgpdf.close()

        output_path = os.path.join(file_name, f"{file_name}.pdf")
        if os.path.exists(output_path):
            os.remove(output_path)
        doc.save(output_path)
        print(f"PDF已保存至: {output_path}")
    finally:
        doc.close()


def decode_base64(s):
    # 解码函数
    m_base64Str = ''
    m_base64Count = 0
    m_END_OF_INPUT = -1
    m_base64Chars_r = [
        'P', 'J', 'K', 'L', 'M', 'N', 'O', 'I',
        '3', 'y', 'x', 'z', '0', '1', '2', 'w',
        'v', 'p', 'r', 'q', 's', 't', 'u', 'o',
        'B', 'H', 'C', 'D', 'E', 'F', 'G', 'A',
        'h', 'n', 'i', 'j', 'k', 'l', 'm', 'g',
        'f', 'Z', 'a', 'b', 'c', 'd', 'e', 'Y',
        'X', 'R', 'S', 'T', 'U', 'V', 'W', 'Q',
        '!', '5', '6', '7', '8', '9', '+', '4'
    ]
    m_reverseBase64Chars = {element: index for index, element in enumerate(m_base64Chars_r)}

    def m_setBase64Str(s):
        nonlocal m_base64Count, m_base64Str
        m_base64Str = s
        m_base64Count = 0

    def m_readReverseBase64():
        nonlocal m_base64Count, m_base64Str, m_reverseBase64Chars
        if not m_base64Str:
            return -1
        while 1:
            if m_base64Count >= len(m_base64Str):
                return -1
            nextCharacter = m_base64Str[m_base64Count]
            m_base64Count += 1
            try:
                if m_reverseBase64Chars[nextCharacter]:
                    return m_reverseBase64Chars[nextCharacter]
            except:
                pass
            if nextCharacter == 'P':
                return 0
        return -1

    def m_ntos(n):
        n = hex(n)
        n = n[2:]
        if len(n) == 1:
            n = "0" + n[-1]
        n = '%' + n
        return html.unescape(n)

    m_setBase64Str(s)
    result = ''
    done = False
    inBuffer = [0, 0, 0, 0]
    inBuffer[0] = m_readReverseBase64()
    inBuffer[1] = m_readReverseBase64()
    while (not done) and (inBuffer[0] != m_END_OF_INPUT) and (inBuffer[1] != m_END_OF_INPUT):
        inBuffer[2] = m_readReverseBase64()
        inBuffer[3] = m_readReverseBase64()
        result += m_ntos((((inBuffer[0] << 2) & 0xff) | inBuffer[1] >> 4))
        if inBuffer[2] != m_END_OF_INPUT:
            result += m_ntos((((inBuffer[1] << 4) & 0xff) | inBuffer[2] >> 2))
            if inBuffer[3] != m_END_OF_INPUT:
                result += m_ntos((((inBuffer[2] << 6) & 0xff) | inBuffer[3]))
            else:
                done = True
        else:
            done = True
        inBuffer[0] = m_readReverseBase64()
        inBuffer[1] = m_readReverseBase64()
    return parse.unquote(result)


if __name__ == "__main__":
    p_code = input('''请输入相应页面的p_code,
如https://m.doc88.com/p-3995949474894.html的p_code为 3995949474894 ：''')
    print('下载中。。。')

    # 第一步获取数据
    try:
        url = f'https://m.doc88.com/doc.php?act=info&p_code={p_code}&key=3854933de90d1dbb321d8ca29eac130a&v=1'
        response = session.get(url, headers=COMMON_HEADERS, timeout=10)
        response.raise_for_status()
        return_data = response.text
    except requests.RequestException as e:
        print(f"获取文档信息失败: {e}")
        exit(1)

    # 第二步解码数据
    try:
        base64str = decode_base64(return_data)
        s = json.loads(base64str)
        gif_host = s['gif_host']
        gif_urls = json.loads(s['gif_struct'])
        file_name = s['name'].replace('/', '_')  # 清理非法文件名
    except (KeyError, json.JSONDecodeError) as e:
        print(f"解析数据失败: {e}")
        exit(1)

    # 第三步下载GIF
    if os.path.exists(file_name):
        print('请先删除同名文件夹')
        exit(1)
    os.mkdir(file_name)

    for index, element in enumerate(gif_urls):
        gif_url = f"{gif_host}/get-{element['u']}.gif"
        filename = f"{index:07d}.gif"  # 文件名补零对齐
        success = download_gif(gif_url, file_name, filename)
        if not success:
            print(f"严重错误: 第 {index + 1} 页下载失败，终止任务")
            exit(1)
        print(f"已下载: {filename}")

    # 第四步合并PDF
    try:
        pic2pdf(file_name)
        print('下载完毕！')
    except Exception as e:
        print(f"合并PDF失败: {e}")
