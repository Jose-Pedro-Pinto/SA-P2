-------------------Modo de Utilização--------------------

-Para selecionar os ficheiros, deve-se clicar nos botões "Selecionar Praças" e em "Selecioar Viagens"
-Para selecionar o número de registos a analisar deve-se preencher a textbox à frente de "Número de registos". 
O valor default é -1, que representa todos os dados.

-Depois de selecionados os ficheiros e depois de se ter escrito o número de registos a analisar, deve-se clicar em "Carregar Dados", 
o que vai carregar os dados dos ficheiros selecionados para as tabelas access("Posturas" e "Viagens" os dados indicados no enunciado
 como ignorar não estão a ser carregados) e depois preencher automaticamente a "ListBox Taxis", que representa a LB1 a "ListBox Praças", 
que presenta a LB2.
Aviso: na versão em ingles(e talvez outras), se o codigo não funcionar é necessario alterar o codigo para fazer replace "." por ",". As duas
versões estão no mesmo codigo, para se trocar comentar e descomentar o codigo nos locais indicados.

-Para usar a Query 1(uma query do access), deve-se clicar em "Query 1", que está na barra esquerda do Access. Depois é pedido um campo com
 a distância mínima, que deve ser preenchido. Vai aparecer uma tabela com os TAXI_ID's e a respetiva distância percorrida
-Para usar a Query 2, deve-se preencher a textbox a seguir a "Query 2: Distancias Totais Maiores que x Valor de X:".
 A seguir devemos clicar no botão "Calcular" que vem a seguir a essa texbox. Os resultados aparecem na listbox em baixo de
"Resultados da Query 2:"
-Para usar a Query 3, deve-se preencher as textbox em baixo de "Distancia > X e mais de Y Viagens" e depois clicar em calcular.
Os resultdos vão aparecer na listbox em baixo de "Resultados da Query 3:"
-Para usar a Query 4, deve-se selecionar TAXI_ID's da listbox "Lista de Ids", preencher a textBox em baixo de "Data de inicio"
com o valor da data de início e  preencher a textBox em baixo de "Data de Fim" com o valor da data de fim. A seguir deve-se
clicar em "Calcular" (que está abaixo da textBox corespondente à "Data de Fim". O resultado vai aparecer em "Gráfio da Query 4"
Nota: datas são do tipo dd-mm-ano, por exemplo 12-2-2013


-------------------Valorização--------------------
Para valorização foi criada uma nova parte no furmulario que calcula o lucro para as pracas de taxis,
para os n taxis com maior lucro e os m taxis com menor lucro (potencialmente negativo, ou seja prejuizo)

-Para calcular o lucro das praças
Preencher o lucro por KM e o custo por minuto e carregar calcular. O resultado aparece na listBox diretamente
a baixo
-Para calcular o lucro do n taxis com maior lucro
Preencher o numero n de taxis(os valores de custo são os mesmos indicados acima). O resultado aparece na 
listBox diretamente a baixo.
-Para calcular o lucro do n taxis com menor lucro
Preencher o numero n de taxis(os valores de custo são os mesmos indicados acima). O resultado aparece na 
listBox diretamente a baixo.
Aviso: na versão em ingles(e talvez outras), se o codigo não funcionar é necessario alterar o codigo para fazer replace "." por ",". As duas
versões estão no mesmo codigo, para se trocar comentar e descomentar o codigo nos locais indicados.


-----------------------Leitura do Código Produzido-------------------------------------------

No código VBA temos 5 módulos. 
-O módulo "ClearData", limpa as tabelas correspondentes aos ficheiros de input.
-O módulo "MathAux" calcula a fórmula de Haversine e outras formulas matematicas necessarias.
-O múdulo "PolyLine" estima a distância percorrida pelo veículo usando a fórmula de Haversine determinada no módulo "MathAux"
-O módulo "ReadData" lê a informação dos ficheiros.
-O módulo "WriteData" escreve a informação lida dos ficheiros em duas tabelas.

Além dos módulos, temos um "Form_Formulário 1" que lida com LB1,LB2 e as querys.
