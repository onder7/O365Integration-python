import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                           QTableWidget, QTableWidgetItem, QTabWidget,
                           QTextEdit, QMessageBox, QProgressBar, QComboBox,
                           QSpinBox, QStatusBar)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QIcon, QFont
from O365 import Account
from datetime import datetime, timedelta
import json
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler
import urllib.parse
import threading
import socket

class OAuth2Handler(BaseHTTPRequestHandler):
    """Handle OAuth2 callback"""
    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-type', 'text/html')
        self.end_headers()
        
        self.server.auth_url = self.path
        
        html = """
        <html><body>
            <h1>Authentication Successful!</h1>
            <p>You can now close this window and return to the application.</p>
            <script>window.close()</script>
        </body></html>
        """
        self.wfile.write(html.encode())
    
    def log_message(self, format, *args):
        pass

class MS365Authenticator:
    """Handle Microsoft 365 Authentication"""
    def __init__(self, client_id, client_secret):
        self.client_id = client_id
        self.client_secret = client_secret
        self.account = None
    
    def get_auth_url(self):
        self.account = Account((self.client_id, self.client_secret))
        return self.account.con.get_authorization_url(
            requested_scopes=['https://graph.microsoft.com/Mail.Read',
                            'https://graph.microsoft.com/Mail.ReadWrite',
                            'offline_access']
        )
    
    def authenticate(self):
        try:
            # Find available port
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(('', 0))
                port = s.getsockname()[1]
            
            # Create callback server
            server = HTTPServer(('localhost', port), OAuth2Handler)
            server.auth_url = None
            
            # Start server thread
            server_thread = threading.Thread(target=server.serve_forever)
            server_thread.daemon = True
            server_thread.start()
            
            # Get and modify auth URL
            auth_url, state = self.get_auth_url()
            redirect_uri = f'http://localhost:{port}'
            auth_url = auth_url.replace('http://localhost', redirect_uri)
            
            # Open browser
            webbrowser.open(auth_url)
            
            # Wait for callback
            while server.auth_url is None:
                pass
            
            # Cleanup
            server.shutdown()
            server.server_close()
            
            # Complete authentication
            auth_url = f"{redirect_uri}{server.auth_url}"
            return self.account.con.request_token(auth_url)
            
        except Exception as e:
            raise Exception(f"Authentication failed: {str(e)}")

class EmailFetchThread(QThread):
    """Background thread for email operations"""
    finished = pyqtSignal(list)
    error = pyqtSignal(str)
    progress = pyqtSignal(int)

    def __init__(self, mail_reader, operation, **kwargs):
        super().__init__()
        self.mail_reader = mail_reader
        self.operation = operation
        self.kwargs = kwargs

    def run(self):
        try:
            if self.operation == "read_inbox":
                results = self.mail_reader.read_inbox(**self.kwargs)
            elif self.operation == "search":
                results = self.mail_reader.search_emails(**self.kwargs)
            elif self.operation == "read_folder":
                results = self.mail_reader.read_folder(**self.kwargs)
            self.finished.emit(results)
        except Exception as e:
            self.error.emit(str(e))

class MS365MailReader:
    """Mail reading functionality"""
    def __init__(self, client_id, client_secret):
        self.client_id = client_id
        self.client_secret = client_secret
        self.account = None
        self.mailbox = None

    def authenticate(self):
        try:
            authenticator = MS365Authenticator(self.client_id, self.client_secret)
            if authenticator.authenticate():
                self.account = authenticator.account
                self.mailbox = self.account.mailbox()
                return True
            return False
        except Exception as e:
            raise Exception(f"Authentication failed: {str(e)}")

    def read_inbox(self, limit=10, days_back=7):
        if not self.mailbox:
            raise Exception("Authentication required!")
            
        query_date = datetime.now() - timedelta(days=days_back)
        messages = self.mailbox.inbox_folder().get_messages(
            limit=limit,
            query=f"receivedDateTime ge {query_date.strftime('%Y-%m-%d')}"
        )
        
        return [self._format_message(msg) for msg in messages]

    def search_emails(self, search_term, folder_name='Inbox', limit=10):
        if not self.mailbox:
            raise Exception("Authentication required!")
            
        if folder_name.lower() == 'inbox':
            folder = self.mailbox.inbox_folder()
        else:
            folder = self.mailbox.get_folder(folder_path=folder_name)
            
        query = f"contains(subject,'{search_term}') or contains(body,'{search_term}')"
        messages = folder.get_messages(limit=limit, query=query)
        
        return [self._format_message(msg) for msg in messages]

    def list_folders(self):
        if not self.mailbox:
            raise Exception("Authentication required!")
        return [folder.name for folder in self.mailbox.list_folders()]

    def _format_message(self, message):
        return {
            'id': message.object_id,
            'subject': message.subject,
            'from': message.sender.address,
            'to': [recipient.address for recipient in message.to],
            'received_time': message.received.strftime('%Y-%m-%d %H:%M:%S'),
            'body': message.body,
            'is_read': message.is_read
        }

class EmailGUI(QMainWindow):
    """Main GUI Window"""
    def __init__(self):
        super().__init__()
        self.mail_reader = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle('MS365 Mail Reader ..::Onder Monder::..')
        self.setGeometry(100, 100, 1200, 800)

        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Authentication section
        auth_layout = QHBoxLayout()
        self.client_id_input = QLineEdit()
        self.client_id_input.setPlaceholderText('Client ID')
        self.client_secret_input = QLineEdit()
        self.client_secret_input.setPlaceholderText('Client Secret')
        self.client_secret_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.auth_button = QPushButton('Authenticate')
        self.auth_button.clicked.connect(self.authenticate)

        auth_layout.addWidget(QLabel('Client ID:'))
        auth_layout.addWidget(self.client_id_input)
        auth_layout.addWidget(QLabel('Client Secret:'))
        auth_layout.addWidget(self.client_secret_input)
        auth_layout.addWidget(self.auth_button)
        layout.addLayout(auth_layout)

        # Tabs
        self.tabs = QTabWidget()
        
        # Inbox tab
        inbox_tab = QWidget()
        inbox_layout = QVBoxLayout(inbox_tab)
        
        inbox_controls = QHBoxLayout()
        self.days_back = QSpinBox()
        self.days_back.setValue(7)
        self.days_back.setMinimum(1)
        self.limit_spin = QSpinBox()
        self.limit_spin.setValue(10)
        self.limit_spin.setMinimum(1)
        self.refresh_button = QPushButton('Refresh Inbox')
        self.refresh_button.clicked.connect(self.refresh_inbox)
        
        inbox_controls.addWidget(QLabel('Days back:'))
        inbox_controls.addWidget(self.days_back)
        inbox_controls.addWidget(QLabel('Limit:'))
        inbox_controls.addWidget(self.limit_spin)
        inbox_controls.addWidget(self.refresh_button)
        inbox_controls.addStretch()
        
        self.emails_table = QTableWidget()
        self.emails_table.setColumnCount(4)
        self.emails_table.setHorizontalHeaderLabels(['Subject', 'From', 'Date', 'Read'])
        self.emails_table.cellClicked.connect(self.show_email_content)
        
        self.email_content = QTextEdit()
        self.email_content.setReadOnly(True)
        
        inbox_layout.addLayout(inbox_controls)
        inbox_layout.addWidget(self.emails_table)
        inbox_layout.addWidget(self.email_content)
        
        # Search tab
        search_tab = QWidget()
        search_layout = QVBoxLayout(search_tab)
        
        search_controls = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText('Search term...')
        self.search_button = QPushButton('Search')
        self.search_button.clicked.connect(self.search_emails)
        self.folder_combo = QComboBox()
        
        search_controls.addWidget(QLabel('Search:'))
        search_controls.addWidget(self.search_input)
        search_controls.addWidget(QLabel('Folder:'))
        search_controls.addWidget(self.folder_combo)
        search_controls.addWidget(self.search_button)
        
        self.search_results = QTableWidget()
        self.search_results.setColumnCount(4)
        self.search_results.setHorizontalHeaderLabels(['Subject', 'From', 'Date', 'Read'])
        self.search_results.cellClicked.connect(self.show_search_content)
        
        self.search_content = QTextEdit()
        self.search_content.setReadOnly(True)
        
        search_layout.addLayout(search_controls)
        search_layout.addWidget(self.search_results)
        search_layout.addWidget(self.search_content)
        
        # Add tabs
        self.tabs.addTab(inbox_tab, "Inbox")
        self.tabs.addTab(search_tab, "Search")
        layout.addWidget(self.tabs)
        
        # Status bar
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        
        # Disable tabs until authentication
        self.tabs.setEnabled(False)

    def authenticate(self):
        client_id = self.client_id_input.text()
        client_secret = self.client_secret_input.text()
        
        if not client_id or not client_secret:
            QMessageBox.warning(self, 'Error', 'Please enter both Client ID and Client Secret')
            return
            
        self.mail_reader = MS365MailReader(client_id, client_secret)
        
        try:
            self.statusBar.showMessage('Authenticating...')
            if self.mail_reader.authenticate():
                self.statusBar.showMessage('Authentication successful')
                self.tabs.setEnabled(True)
                self.folder_combo.addItems(['Inbox'] + self.mail_reader.list_folders())
                self.refresh_inbox()
            else:
                QMessageBox.warning(self, 'Error', 'Authentication failed')
                self.statusBar.showMessage('Authentication failed')
        except Exception as e:
            QMessageBox.critical(self, 'Error', str(e))
            self.statusBar.showMessage('Authentication error')

    def refresh_inbox(self):
        if not self.mail_reader:
            return
            
        self.statusBar.showMessage('Fetching emails...')
        self.thread = EmailFetchThread(
            self.mail_reader,
            "read_inbox",
            limit=self.limit_spin.value(),
            days_back=self.days_back.value()
        )
        self.thread.finished.connect(self.display_emails)
        self.thread.error.connect(self.show_error)
        self.thread.start()

    def search_emails(self):
        if not self.mail_reader:
            return
            
        search_term = self.search_input.text()
        if not search_term:
            QMessageBox.warning(self, 'Error', 'Please enter a search term')
            return
            
        self.statusBar.showMessage('Searching...')
        self.thread = EmailFetchThread(
            self.mail_reader,
            "search",
            search_term=search_term,
            folder_name=self.folder_combo.currentText(),
            limit=self.limit_spin.value()
        )
        self.thread.finished.connect(self.display_search_results)
        self.thread.error.connect(self.show_error)
        self.thread.start()

    def display_emails(self, emails):
        self.emails_table.setRowCount(len(emails))
        self._fill_table(self.emails_table, emails)
        self.statusBar.showMessage('Emails loaded')

    def display_search_results(self, emails):
        self.search_results.setRowCount(len(emails))
        self._fill_table(self.search_results, emails)
        self.statusBar.showMessage('Search completed')

    def _fill_table(self, table, emails):
        for i, email in enumerate(emails):
            table.setItem(i, 0, QTableWidgetItem(email['subject']))
            table.setItem(i, 1, QTableWidgetItem(email['from']))
            table.setItem(i, 2, QTableWidgetItem(email['received_time']))
            table.setItem(i, 3, QTableWidgetItem('Yes' if email['is_read'] else 'No'))
        
        # Resize columns to content
        table.resizeColumnsToContents()

    def show_email_content(self, row, col):
        email_subject = self.emails_table.item(row, 0).text()
        email_from = self.emails_table.item(row, 1).text()
        email_date = self.emails_table.item(row, 2).text()
        
        content = f"Subject: {email_subject}\nFrom: {email_from}\nDate: {email_date}\n\n"
        self.email_content.setText(content)

    
    def show_search_content(self, row, col):
        email_subject = self.search_results.item(row, 0).text()
        email_from = self.search_results.item(row, 1).text()
        email_date = self.search_results.item(row, 2).text()
        
        content = f"Subject: {email_subject}\nFrom: {email_from}\nDate: {email_date}\n\n"
        self.search_content.setText(content)

    def show_error(self, error_message):
        QMessageBox.critical(self, 'Error', error_message)
        self.statusBar.showMessage('Error occurred')

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # Modern look
    ex = EmailGUI()
    ex.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()