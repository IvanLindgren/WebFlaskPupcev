from flask import Flask, request, render_template
import pandas as pd
from ldap3 import Server, Connection, ALL, MODIFY_REPLACE

app = Flask(__name__)

LDAP_SERVER = ''
LDAP_USER = ''
LDAP_PASSWORD = ''
BASE_DN = ''


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part"

    file = request.files['file']

    if file.filename == '':
        return "No selected file"

    if file:
        df = pd.read_excel(file)
        for index, row in df.iterrows():
            sync_user(row['номер студенческого'], row['ФИО студента'], row['баллы'])

        return "Синхронизация аккаунтов прошла успешна..."

    return "File processing error"


def sync_user(student_id, full_name, points):
    server = Server(LDAP_SERVER, get_info=ALL)
    conn = Connection(server, LDAP_USER, LDAP_PASSWORD, auto_bind=True)

    search_filter = f'(uid={student_id})'
    conn.search(BASE_DN, search_filter, attributes=['uid', 'cn', 'points'])

    if conn.entries:
        # Если пользователь есть - обновим баллы:
        dn = conn.entries[0].entry_dn
        conn.modify(dn, {'points': [(MODIFY_REPLACE, [points])]})
    else:
        # Если пользователя нет - зарегистрируем и обновим банны:
        dn = f'uid={student_id},{BASE_DN}'
        conn.add(dn, ['inetOrgPerson'], {'uid': student_id, 'cn': full_name, 'points': points})

    conn.unbind()


if __name__ == '__main__':
    app.run(debug=True)
