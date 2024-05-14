from waitress import serve

from app import app
import os, logging, socket

# 
def deploy():
    PORT = 5000
    os.system('cls')
    logger = logging.getLogger('waitress')
    logger.setLevel(logging.WARNING)
    print(f'Server Started | http://{socket.gethostname()}:{PORT}')
    print('Running...')
    serve(app, host='0.0.0.0', port=PORT, url_scheme='https')

if __name__ == '__main__':
    deploy()