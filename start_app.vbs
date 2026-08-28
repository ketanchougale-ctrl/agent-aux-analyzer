Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c cd /d ""C:\Users\ketanchougale\CascadeProjects\agent-aux-analyzer"" && streamlit run app.py --server.port 8501 --server.headless true", 0, False
