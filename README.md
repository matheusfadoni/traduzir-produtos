Script feito compativel com python 3.10 ou superior
Recomendo antes de tudo usar o script split-planilha.py
Ele separa a planilha principal em arquivos ".xlsx" de 500 em 500 linhas 
Pois qualquer problema com conexão ou outras enventualidade pode interromper a execução do código e perder todo o trabalho de tradução realizado antes de salvar o arquivo final.
As planilhas com tradução devem ser revisadas após tradução pois google translator muitas vezes retorna traduções muito erradas.
O formato da planilha usado foi:
Coluna A: código do produto (SKU)
Coluna B: Titulo do produto em espanhol "esp"
Coluna C: Tradução em Português "trad" (esta coluna deve ser deixada vazia para o script salvar a tradução aqui)

na variável nome_arquivo configure o titulo do seu arquivo sem a extensão e em file_path configure a extensão de planilha usada
exemplo:
  nome_arquivo = 'Parte12'
  file_path = nome_arquivo + '.xlsx'
se tudo correr bem, seguindo este exemplo, o arquivo de saída se chamará "Parte12-revisar.xlsx"
você pode configurar isto no arquivo também alterando a variável "output_file"
Você deve alterar o script para ajustar ás suas necessidades está muito bem comentado para facilitar o entendimento.
Para executar a tratução abra o terminal vá até a pasta do script e digite "python traduzir-produtos-google-v2.py"
