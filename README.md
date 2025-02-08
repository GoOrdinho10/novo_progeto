Fluxograma do Projeto de Captura de 
Dados Climáticos 

1. Início 
● Executar o programa 

2. Captura de Dados Climáticos 
● Abrir o navegador via Selenium 
● Acessar a página do Google para busca de clima 
● Aguardar o carregamento do widget do tempo 
● Capturar temperatura, umidade, vento e chuva 
● Fechar o navegador 

3. Verificação da Captura de Dados 
● Se bem-sucedida, seguir para salvar os dados na planilha 
● Se mal-sucedida, exibir erro e encerrar 

4. Salvamento dos Dados na Planilha 
● Verificar se o arquivo Excel existe: 
○ Se sim, abrir e adicionar nova linha 
○ Se não, criar novo arquivo e adicionar cabeçalho 
● Inserir os dados capturados com data e hora 
● Salvar o arquivo 

5. Verificação do Salvamento 
● Se bem-sucedido, atualizar a interface com os dados capturados 
● Se falha, exibir mensagem de erro 

6. Interface Gráfica (Tkinter) 
● Criar janela 
● Criar botão para captura de dados 
● Exibir os dados capturados na tela 

7. Encerramento 
● Finalizar execução 
