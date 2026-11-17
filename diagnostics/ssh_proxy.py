"""沿用本機既有 HTTP 代理，僅連接 GitHub SSH 443；不輸出憑據。"""
import base64,os,socket,ssl,sys,threading,urllib.parse,urllib.request

def main():
    value=urllib.request.getproxies().get('https') or urllib.request.getproxies().get('http')
    if not value:raise RuntimeError('本機沒有既有代理')
    proxy=urllib.parse.urlsplit(value if '://' in value else 'http://'+value)
    if proxy.scheme not in {'http','https'} or not proxy.hostname:raise ValueError('未支援代理類型')
    connection=socket.create_connection((proxy.hostname,proxy.port or (443 if proxy.scheme=='https' else 80)),timeout=15)
    try:
        if proxy.scheme=='https':connection=ssl.create_default_context().wrap_socket(connection,server_hostname=proxy.hostname)
        headers=['CONNECT ssh.github.com:443 HTTP/1.1','Host: ssh.github.com:443']
        if proxy.username is not None:
            secret=urllib.parse.unquote(proxy.username)+':'+urllib.parse.unquote(proxy.password or '')
            headers.append('Proxy-Authorization: Basic '+base64.b64encode(secret.encode()).decode())
        connection.sendall(('\r\n'.join(headers)+'\r\n\r\n').encode('ascii'))
        response=bytearray()
        while not response.endswith(b'\r\n\r\n'):
            part=connection.recv(1)
            if not part or len(response)>=65536:raise RuntimeError('代理握手未完成')
            response.extend(part)
        if response.split(b'\r\n',1)[0].split()[1]!=b'200':raise RuntimeError('代理拒絕連線')
        connection.settimeout(None)
        def upload():
            try:
                while True:
                    raw=os.read(sys.stdin.fileno(),65536)
                    if not raw:break
                    connection.sendall(raw)
                connection.shutdown(socket.SHUT_WR)
            except OSError:pass
        threading.Thread(target=upload,daemon=True).start()
        while True:
            raw=connection.recv(65536)
            if not raw:break
            sys.stdout.buffer.write(raw);sys.stdout.buffer.flush()
    finally:connection.close()

if __name__=='__main__':
    try:main()
    except Exception:sys.stderr.write('Company GitHub proxy connection failed.\n');raise SystemExit(2)
