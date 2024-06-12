# Automatização de Dayle

Desenvolvi com python uma automatização que criação de um excel para a Dayle.
O programa inicia excluido o arquivo extraido do dia anterio e inifica a interface grafica, a qual tem a responsabilidade de receber login e senha do usuario do zabbix. usuario e senha são usados para efetuar login no zabbix, o usuario é utulizado durante todo o programa, 
ele é passado como argumento das funções, é usado para verificar o caminho de cada maquina visto que para cada usuario o caminho devera alterar com base no usuario.
Login no zabbix é realizado com auxilio da biblioteca selenium, ele abre o navegador para realizar login no zabbix, na pagina aberta é selecionado os elementos pelo ID e realizado o click em login, com isso abre a tela de login onde utilizamos o usuario e senha que pegamos 
anteriomente, com isso finalizar a parte do login, onde chamamos a função de extração onde usei a biblioteca pyautogui (que me deu a ideia de realizar este projeto), e realizo a filtragem dos dados necessario para a extração após isso realizo a extração. 
Com a extração dos dados realizada o programa começa a formatar o excel com exclusão, renomeação e ordenação das colunar. o segundo passo agora é a estilização do excel que é realizado com a biblioteca openpyxl, onde formatamos como tabela todas as cedulas necessaria para
a tabela.

