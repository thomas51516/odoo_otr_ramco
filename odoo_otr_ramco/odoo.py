import paramiko
client = paramiko.SSHClient()
client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
client.connect(hostname='41.207.181.214',username='ramco', password='PwG.2Kk:r648Hx', port=8202, allow_agent=False, look_for_keys=False)
sftp = client.open_sftp()

sftp.put('fichier_text.xml','fichier_text.xml')
