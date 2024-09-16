from flask import Flask, request, render_template
import pandas as pd
from ldap3 import Server, Connection, ALL, MODIFY_REPLACE

app = Flask(__name__)

LDAP_SERVER = 'ldap://localhost'
LDAP_USER = 'cn=Manager,dc=my-domain,dc=com'
LDAP_PASSWORD = 'secret'
BASE_DN = 'dc=my-domain,dc=com'


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "Вы не загрузили файл"

    file = request.files['file']

    if file.filename == '':
        return "Файл не выбран"

    if file:
        df = pd.read_excel(file)
        for index, row in df.iterrows():
            print(f"Обработка пользователя: {row[1]}, {row[2]}, {row[3]}")
            sync_user(row[1], row[2], row[3])

        return "Аккаунты были загружены"

    return "Ошибка при обработке файла"


def sync_user(student_id, full_name, points):
    server = Server(LDAP_SERVER, get_info=ALL)
    conn = Connection(server, LDAP_USER, LDAP_PASSWORD, auto_bind=True)

    search_filter = f'(uid={student_id})'
    conn.search(BASE_DN, search_filter, attributes=['uid', 'cn', 'description'])

    if conn.entries:
        # Обновляем баллы, если пользователь уже существует
        dn = conn.entries[0].entry_dn
        print(f"Updating user: {dn}")
        conn.modify(dn, {'description': [(MODIFY_REPLACE, [str(points)])]})
    else:
        # Если нет пользователя - создаем нового
        dn = f'uid={student_id},ou=People,{BASE_DN}'
        print(f"Creating user: {dn}")


        if isinstance(full_name, str):
            sn = full_name.split()[1] if len(full_name.split()) > 1 else full_name
        else:
            sn = "Unknown"

        conn.add(dn, ['inetOrgPerson'], {'uid': student_id, 'cn': full_name, 'sn': sn, 'description': str(points)})

    conn.unbind()


if __name__ == '__main__':
    app.run(debug=True)
