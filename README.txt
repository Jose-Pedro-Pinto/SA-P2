-------------------Modo de Utiliza��o--------------------

-Para selecionar os ficheiros, deve-se clicar nos bot�es "Selecionar Pra�as" e em "Selecioar Viagens"
-Para selecionar o n�mero de registos a analisar deve-se preencher a textbox � frente de "N�mero de registos". 
O valor default � -1, que representa todos os dados.

-Depois de selecionados os ficheiros e depois de se ter escrito o n�mero de registos a analisar, deve-se clicar em "Carregar Dados", 
o que vai carregar os dados dos ficheiros selecionados para as tabelas access("Posturas" e "Viagens" os dados indicados no enunciado
 como ignorar n�o est�o a ser carregados) e depois preencher automaticamente a "ListBox Taxis", que representa a LB1 a "ListBox Pra�as", 
que presenta a LB2.
Aviso: na vers�o em ingles(e talvez outras), se o codigo n�o funcionar � necessario alterar o codigo para fazer replace "." por ",". As duas
vers�es est�o no mesmo codigo, para se trocar comentar e descomentar o codigo nos locais indicados.

-Para usar a Query 1(uma query do access), deve-se clicar em "Query 1", que est� na barra esquerda do Access. Depois � pedido um campo com
 a dist�ncia m�nima, que deve ser preenchido. Vai aparecer uma tabela com os TAXI_ID's e a respetiva dist�ncia percorrida
-Para usar a Query 2, deve-se preencher a textbox a seguir a "Query 2: Distancias Totais Maiores que x Valor de X:".
 A seguir devemos clicar no bot�o "Calcular" que vem a seguir a essa texbox. Os resultados aparecem na listbox em baixo de
"Resultados da Query 2:"
-Para usar a Query 3, deve-se preencher as textbox em baixo de "Distancia > X e mais de Y Viagens" e depois clicar em calcular.
Os resultdos v�o aparecer na listbox em baixo de "Resultados da Query 3:"
-Para usar a Query 4, deve-se selecionar TAXI_ID's da listbox "Lista de Ids", preencher a textBox em baixo de "Data de inicio"
com o valor da data de in�cio e  preencher a textBox em baixo de "Data de Fim" com o valor da data de fim. A seguir deve-se
clicar em "Calcular" (que est� abaixo da textBox corespondente � "Data de Fim". O resultado vai aparecer em "Gr�fio da Query 4"
Nota: datas s�o do tipo dd-mm-ano, por exemplo 12-2-2013


-------------------Valoriza��o--------------------
Para valoriza��o foi criada uma nova parte no furmulario que calcula o lucro para as pracas de taxis,
para os n taxis com maior lucro e os m taxis com menor lucro (potencialmente negativo, ou seja prejuizo)

-Para calcular o lucro das pra�as
Preencher o lucro por KM e o custo por minuto e carregar calcular. O resultado aparece na listBox diretamente
a baixo
-Para calcular o lucro do n taxis com maior lucro
Preencher o numero n de taxis(os valores de custo s�o os mesmos indicados acima). O resultado aparece na 
listBox diretamente a baixo.
-Para calcular o lucro do n taxis com menor lucro
Preencher o numero n de taxis(os valores de custo s�o os mesmos indicados acima). O resultado aparece na 
listBox diretamente a baixo.
Aviso: na vers�o em ingles(e talvez outras), se o codigo n�o funcionar � necessario alterar o codigo para fazer replace "." por ",". As duas
vers�es est�o no mesmo codigo, para se trocar comentar e descomentar o codigo nos locais indicados.


-----------------------Leitura do C�digo Produzido-------------------------------------------

No c�digo VBA temos 5 m�dulos. 
-O m�dulo "ClearData", limpa as tabelas correspondentes aos ficheiros de input.
-O m�dulo "MathAux" calcula a f�rmula de Haversine e outras formulas matematicas necessarias.
-O m�dulo "PolyLine" estima a dist�ncia percorrida pelo ve�culo usando a f�rmula de Haversine determinada no m�dulo "MathAux"
-O m�dulo "ReadData" l� a informa��o dos ficheiros.
-O m�dulo "WriteData" escreve a informa��o lida dos ficheiros em duas tabelas.

Al�m dos m�dulos, temos um "Form_Formul�rio 1" que lida com LB1,LB2 e as querys.
