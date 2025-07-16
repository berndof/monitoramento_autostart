$edge = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"


Start-Process $edge -ArgumentList "--app=https://monitoramento.scc.com.br/zabbix.php?action=dashboard.view --start-fullscreen"


Start-Process $edge -ArgumentList "--app=https://rotadigital.scc.com.br/d/fejvra6ovbbwga/ti-lages-links?orgId=1&from=now-1h&to=now&timezone=browser&refresh=1m"