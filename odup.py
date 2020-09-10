
import json
import os
import random

import click
import requests

USER_AGENTS = [
    "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; "
    "SV1; AcooBrowser; .NET CLR 1.1.4322; .NET CLR 2.0.50727)",
    "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0; Acoo Browser; "
    "SLCC1; .NET CLR 2.0.50727; Media Center PC 5.0; .NET CLR 3.0.04506)",
    "Mozilla/4.0 (compatible; MSIE 7.0; AOL 9.5; AOLBuild 4337.35; "
    "Windows NT 5.1; .NET CLR 1.1.4322; .NET CLR 2.0.50727)",
    "Mozilla/5.0 (Windows; U; MSIE 9.0; Windows NT 9.0; en-US)",
    "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Win64; x64; "
    "Trident/5.0; .NET CLR 3.5.30729; .NET CLR 3.0.30729; "
    ".NET CLR 2.0.50727; Media Center PC 6.0)",
    "Mozilla/5.0 (compatible; MSIE 8.0; Windows NT 6.0; "
    "Trident/4.0; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; "
    ".NET CLR 3.5.30729; .NET CLR 3.0.30729; "
    ".NET CLR 1.0.3705; .NET CLR 1.1.4322)",
    "Mozilla/4.0 (compatible; MSIE 7.0b; "
    "Windows NT 5.2; .NET CLR 1.1.4322; .NET CLR 2.0.50727; "
    "InfoPath.2; .NET CLR 3.0.04506.30)",
    "Mozilla/5.0 (Windows; U; Windows NT 5.1; zh-CN) "
    "AppleWebKit/523.15 (KHTML, like Gecko, Safari/419.3) "
    "Arora/0.3 (Change: 287 c9dfb30)",
    "Mozilla/5.0 (X11; U; Linux; en-US) AppleWebKit/527+ "
    "(KHTML, like Gecko, Safari/419.3) Arora/0.6",
    "Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; "
    "rv:1.8.1.2pre) Gecko/20070215 K-Ninja/2.1.1",
    "Mozilla/5.0 (Windows; U; Windows NT 5.1; zh-CN; rv:1.9) "
    "Gecko/20080705 Firefox/3.0 Kapiko/3.0",
    "Mozilla/5.0 (X11; Linux i686; U;) "
    "Gecko/20070322 Kazehakase/0.4.5",
    "Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.8) "
    "Gecko Fedora/1.9.0.8-1.fc10 Kazehakase/0.5.6",
    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/535.11 "
    "(KHTML, like Gecko) Chrome/17.0.963.56 Safari/535.11",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_3) "
    "AppleWebKit/535.20 (KHTML, like Gecko) Chrome/19.0.1036.7 Safari/535.20",
    "Opera/9.80 (Macintosh; Intel Mac OS X 10.6.8; U; fr) "
    "Presto/2.9.168 Version/11.52"
]


def parseConf():
    with open("conf.json", "rb+") as f:
        return json.loads(f.read())


def parseUrl(url):
    tenant = url.split('/')[2]
    mail = url.split('/')[6]
    return tenant, mail


def getCookies(url):
    headers = {
        'User-Agent': random.choice(USER_AGENTS)
    }
    response = requests.get(url, headers=headers)
    cookies = response.cookies.get_dict()
    # print('提取 FedAuth:' + cookies['FedAuth'])
    return cookies


def getAccessToken(url):
    tenant, mail = parseUrl(url)
    cookies = getCookies(url)
    url = "https://" + tenant + "/personal/" + mail + \
        "/_api/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream?@a1='/personal/" + mail + \
        "/Documents'&RootFolder=/personal/" + mail + \
        "/Documents/&TryNewExperienceSingle=TRUE"

    headers = {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'User-Agent': random.choice(USER_AGENTS)
    }

    payload = {
        "parameters": {
            "__metadata": {"type": "SP.RenderListDataParameters"},
            "RenderOptions": 136967,
            "AllowMultipleValueFilterForTaxonomyFields": True,
            "AddRequiredFields": True
        }
    }

    response = requests.post(url, cookies=cookies,
                             headers=headers, data=json.dumps(payload))

    payload = json.loads(response.text)
    token = payload['ListSchema']['.driveAccessToken'][13:]
    api_url = payload['ListSchema']['.driveUrl'] + '/'
    shared_folder = payload['ListData']['Row'][0]['FileRef'].split('/')[-1]
    # print('提取 目录名:' + shared_folder)
    # print('提取 AccessToken:' + token)
    # print('提取 api_url:' + api_url)
    return token, api_url, shared_folder


def HRS(size, precision=2):
    suffixes = ['B', 'KB', 'MB', 'GB', 'TB']
    suffixIndex = 0
    while size > 1024 and suffixIndex < 4:
        suffixIndex += 1  # increment the index of the suffix
        size = size/1024.0  # apply the division
    return "%.*f%s" % (precision, size, suffixes[suffixIndex])


@click.group()
def cli():
    pass


@click.command()
@click.option("--share", default='', help='分享链接')
def init(share):
    """ 
    初始化配置文件
    """
    if share == '':
        click.echo('Error: 请输入分享地址')
        os._exit(0)
    data = {'shareLink': share}
    json_str = json.dumps(data)
    with open('conf.json', 'w') as f:
        f.write(json_str)
    click.echo('[conf.json] 配置文件创建成功')


@click.command()
@click.option("--file", default='', help='需要上传的文件')
@click.option("--path", default='', help='上传路径')
def upload(file, path):
    """ 
    OneDrive 文件上传工具
    """
    if file == '' or path == '':
        click.echo('Error: 请输入需要上传的文件路径和上传目标路径')
        os._exit(0)
    conf = []
    try:
        conf = parseConf()
    except Exception:
        click.echo('Error: [conf.json] 配置文件未找到或不存在')
        os._exit(0)
    shareLink = conf['shareLink']
    token, api_url, shared_folder = getAccessToken(shareLink)
    path = path.strip('/') + '/'
    file_size = os.path.getsize(file)
    # print('文件大小:' + str(round(float(file_size / 1024 / 1024), 2)) + ' MB')
    (filepath, tempfilename) = os.path.split(file)
    uploadpath = (path + tempfilename).strip('/')
    click.echo(f'开始上传（{HRS(file_size)}），上传地址：{uploadpath}')
    if file_size < 1024 * 1024 * 4:
        headers = {
            'Authorization': 'Bearer ' + token,
            'Content-Type': 'application/json'
        }
        f = open(file, "rb")
        url = f'{api_url}items/root:/{shared_folder}/{uploadpath}:/content'
        response = requests.put(url, headers=headers,
                                data=f.read())
        file_id = json.loads(response.text)['id']
        # print('提取 文件ID:' + file_id)
    else:
        # 分片上传最大15G
        headers = {
            'Authorization': 'Bearer ' + token,
            'Content-Type': 'application/json'
        }
        url = f'{api_url}items/root:/{shared_folder}/{uploadpath}:/createUploadSession'
        response = requests.post(url, headers=headers)
        uploadUrl = json.loads(response.text)['uploadUrl']

        with click.progressbar(label='File Uploading', length=file_size) as bar, open(file, "rb") as f:
            while True:
                data = f.read(20 * 1024 * 1024)
                if not data:
                    file_id = json.loads(response.text)['id']
                    # print('提取 文件ID:' + file_id)
                    bar.finish()
                    break
                headers = {
                    'Content-Length': str(len(data)),
                    'Content-Range': 'bytes ' + str(f.tell() - len(data)) + '-' + str(f.tell() - 1) + '/' + str(
                        file_size)
                }
                bar.update(len(data))
                response = requests.put(
                    uploadUrl, headers=headers, data=data)
    headers = {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
    }
    url = api_url + 'items/' + file_id + '/content'
    response = requests.get(url, headers=headers, allow_redirects=False)
    download_link = response.headers['Location']
    print(f'文件上传成功，下载直链地址：\n{download_link}')


cli.add_command(upload)
cli.add_command(init)

if __name__ == '__main__':
    cli()
