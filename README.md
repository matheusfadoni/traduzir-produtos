Script feito compativel com python 3.10 ou superior<br>
<br>
Recomendo antes de tudo usar o script split-planilha.py<br>
Ele separa a planilha principal em arquivos ".xlsx" de 500 em 500 linhas <br>
Pois qualquer problema com conexão ou outras enventualidade pode interromper a execução do código e perder todo o trabalho de tradução realizado antes de salvar o arquivo final.<br>
As planilhas com tradução devem ser revisadas após tradução pois google translator muitas vezes retorna traduções muito erradas.<br>
O formato da planilha usado foi:<br>
Coluna A: código do produto (SKU)<br>
Coluna B: Titulo do produto em espanhol "esp"<br>
Coluna C: Tradução em Português "trad" (esta coluna deve ser deixada vazia para o script salvar a tradução aqui)<br>
<br>
na variável nome_arquivo configure o titulo do seu arquivo sem a extensão e em file_path configure a extensão de planilha usada<br>
exemplo:<br>
  nome_arquivo = 'Parte12'<br>
  file_path = nome_arquivo + '.xlsx'<br>
se tudo correr bem, seguindo este exemplo, o arquivo de saída se chamará "Parte12-revisar.xlsx"<br>
você pode configurar isto no arquivo também alterando a variável "output_file"<br>
Você deve alterar o script para ajustar ás suas necessidades está muito bem comentado para facilitar o entendimento.<br>
Para executar o split das planilhas ou a tradução abra o terminal vá até a pasta do script e digite: python + nome_do_arquivo.py<br>
exemplo: "python traduzir-produtos-google-v2.py"<br>
<br>
