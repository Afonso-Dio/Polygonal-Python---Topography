import matplotlib.pyplot as plt
import numpy as np
import math as m
import lat_lon_parser as llp
import os
from pandas import DataFrame
import xlrd
import pandas as pd

#Seção 1 : Conversão de graus para graus decimais em python para topografia
    #||Formato de inserção de dados de angulos == xxx°xx°xx.xx°///Exemplo: 2°22°33.56565°||
        #Dicionario de siglas usadas no código:
            #Pe = Precisão do equipamento
            #Pn = Pontos
            #arraySalvar = aS
            #Ang = angulo
            #AngH = angulo horizontal
            #gDC = Graus decimais
            #aSAngHDP = array salvar angulos horizontais decimais por ponto
            #aSAngHCF = array salvar angulos horizontais calculator format
            #Ta = Tolerancia angular
            #Tl = Tolerancia linear
            #Oae = Orientação do erro angular
            #Ea = Erro angular
            #EaU = Erro angular p/ usuario
            #TaU = tolerancia angular p/ usuario
            #aSAngHC = array salvar angulo horizontal corrigido
            #aSAngHCU = array salvar angulo horizontal corrigido para usuário
            #mcea = método de compensação do erro angular
            #aSAzU = array salvar azimutes p/ usuário
            #xPs = xPontos;yPs = yPontos
            #CxPs = Corrigidos xPontos;CyPs = Corrigidos yPontos
            #Dhs = Distancias Horizontais
            #xPPi e yPPi = Variavel usada no calculo de coordenadas provisórias; 
            #xPPci e yPPci = Variavel usada no calculo de coordenadas corrigidas;

#Seção 1.1 - Usando condições(True/False) para evitar a inserção de dados incorretos em Pn(Número de pontos da poligonal) e inserção de dados inicais do calculo
#A flag serve montar um processo anti-quebra do código, que pode ser usado de duas formas como na seção 1.2 #Essa vetor guarda dados para limitar a entrada no input de angulos 
#visto que os valores correspondentes possuem essa limitação#array salvar angulos horizontais decimais por ponto...Ex:[18.3543453°,12.3984u2839°,....] 
#array salvar angulos horizontais calculator format....Ex: 18°10'32" 
#
#
#
#
#
criarprojeto=str(input('Novo projeto? (Y) para sim, (N) para entrar em um existente: '))
def start(): #função de start, ve se a pasta projetos está criada e se nao estiver ela cria
    if not os.path.exists('Projetos'):
        os.mkdir('Projetos')
        os.chdir('Projetos');
def crpj(criarprojeto): #função para criar projeto caso a opção de criar um seja escolhida
    if criarprojeto == 'Y':
        os.chdir('Projetos');
        sel=str(input('Insira o nome do projeto: '))
        os.mkdir(sel);
        print('Novo projeto [{}] criado com sucesso'.format(sel))
        os.chdir(sel)
        return poligonal(sel)
    if criarprojeto == 'N':
        select()
def select(): #função de seleção de projetos ja criados
    a=os.listdir('Projetos')
    sel=str(input('Selecione um dos projetos {}: '.format(a)))
    if sel == 'Criar':
        criarprojeto = 'Y'
        return crpj(criarprojeto)
    for i in range(len(a)):
        if sel == a[i]:
            sel1=input('Projeto [{}] selecionado, deseja confirmar?(Y) ou voltar a seleção?(N): '.format(sel))
            if sel1 == 'N':
                select()
            if sel1 == 'Y':
                print('Projeto [{}] aberto'.format(sel))
                poligonal(sel)
            if not sel1 == 'Y' and sel1 == 'N':
                print('Selecione uma opção valida')
                select()
    if not os.path.exists(sel):
        print('Selecione um projeto existente ou crie outro, inserindo (Criar)')
        select()

def poligonal(sel):
    dir=str(sel)
    diretório=str('C:/Users/Dio/Documents/visual code/Projetos/'+dir+'/Projeto'+dir+'.xlsx')

    flag=True;n=[360,60,60];aSAngHDP=[];aSAngHCF=[];aSAngAz=[] 
    while True:
            try: 
                Pe = input('Insira a precisão do equipamento: ');Pe = Pe.split("\u00B0");Pe=(int(Pe[0]),int(Pe[1]),float(Pe[2]));Pe=llp.to_dec_deg(Pe[0],Pe[1],Pe[2]);PeU1=llp.to_deg_min_sec(Pe)
                PeU='{}\u00B0{}\u02BC{}\u02EE'.format(int(PeU1[0]),int(PeU1[1]),float(PeU1[2]))
                if not(PeU1[0] < n[0] and PeU1[1] < n[1] and PeU1[2] < n[2]):                  #Este if cria a condição de escolha para o programa identificar que o valor inserado está errado
                    raise ValueError('Valor fora do intervalo permitido')
                break
            except ValueError as error:
                print(error)
    print('------------------------------------------------------------------------------------------------------------------------------------------')
    #
    #
    #
    #
    #
    while (flag):
        try:
            Pn=int(input('Insira o numero de pontos da poligonal:'));
            if(type(Pn) is int):
                flag=False
        except:
            print('Insira um valor inteiro')
    print('------------------------------------------------------------------------------------------------------------------------------------------')
    #O primeiro while está vinculado à evitar a quebra do código, fazendo com que sempre que o usuario digite um valor diferente do solicitado o programa retorne o mesmo input

    #
    #
    #
    #
    #
    #Seção 1.2 - Usando uma versão diferente de condições(True/False) para evitar a inserção de angulos incorretos e fazer os devidos calculos de conversão para as seguintes seções

    for i in range(Pn):
        while True:
            try: 
                AngH = input('Insira o angulo horizontal do ponto em dados de graus, minutos e segundos, por ordem de pontos: ');AngH = AngH.split("\u00B0");AngH = (int(AngH[0]),int(AngH[1]),float(AngH[2]))                #Entrada dos angulos horizontais #Método de separação de variaveis #Separando os elementos por vetor 
                gDC = llp.to_dec_deg(AngH[0],AngH[1],AngH[2])                                         #Formato de aparencia pro usuário, comumente escrito nos exercicios a mão #Conversão de graus, minutos e segundos para graus decimais        
                AngsHU=llp.to_deg_min_sec(gDC);AngsHU='{}\u00B0{}\u02BC{}\u02EE'.format(int(AngsHU[0]),int(AngsHU[1]),float(AngsHU[2]))
                if not(AngH[0] < n[0] and AngH[1] < n[1] and AngH[2] < n[2] and type(AngH[0]) is int and type(AngH[1]) is int and type(AngH[2]) is float):                  #Este if cria a condição de escolha para o programa identificar que o valor inserado está errado
                    raise ValueError('Valor fora do intervalo permitido')
                break
            except ValueError as error:
                print(error)
        aSAngHCF.append(AngsHU);aSAngHDP.append(gDC)
    Ta=float(Pe*m.sqrt(Pn))                                         #Tolerancia angular, usando comando da biblioteca math pra fazer raiz quadrada, abaixo TaU
    TaU=llp.to_deg_min_sec(Ta);TaU='{}\u00B0{}\u02BC{}\u02EE'.format(int(TaU[0]),int(TaU[1]),float(TaU[2]));
    aSAngHDP=[float(x) for x in aSAngHDP]                    #Comando para transformar todos os dados da linha em float
    print('------------------------------------------------------------------------------------------------------------------------------------------')

    #
    #
    #
    #
    #
    #Seção 1.3 - Seleção de angulos internos ou externos e calculo do erro angular com verificação de tolerancia

    while True:
        try:
            Oae=input('Os angulos horizontais são internos ou externos? Digite (1) para internos e (2) para externos: ')
            if not(Oae == '1' or Oae == '2'):
                raise ValueError('Selecione a opção (1) ou (2)')
            break
        except ValueError as error:
            print(error)
    #
    #
    #
    if Oae == '1':   #Nestes dois 'if' é usado o procedimento de calculo do erro angular de forma que varie conforme for angulos externos ou internos,  
        Ea=(sum(aSAngHDP))-((Pn-2)*180);
        EaU=llp.to_deg_min_sec(Ea);EaUa=round(EaU[2],3);Eac=llp.to_dec_deg(EaU[0],EaU[1],EaUa);EaU=llp.to_deg_min_sec(Eac);
        EaU='{}\u00B0{}\u02BC{}\u02EE'.format(int(EaU[0]),int(EaU[1]),float(EaU[2]));EaAbs=abs(Ea)
    #
    #
    #
    if Oae == '2':
        Ea=(sum(aSAngHDP))-((Pn+2)*180);
        EaU=llp.to_deg_min_sec(Ea);EaUa=round(EaU[2],3);Eac=llp.to_dec_deg(EaU[0],EaU[1],EaUa);EaU=llp.to_deg_min_sec(Eac);
        EaU='{}\u00B0{}\u02BC{}\u02EE'.format(int(EaU[0]),int(EaU[1]),float(EaU[2]));EaAbs=abs(Ea)
    #
    #
    #
    if EaAbs < Ta:  #Este 'if' verifica se o erro angular está dentro da tolerancia, o comando 'abs' é para colocar o valor dentro dos parenteses em modulo deixando ele sempre positivo
        print('O erro angular = {} está dentro da tolerancia = {}'.format(EaU,TaU)) #usando format para modelar a mensagem para o usuário
    #
    #
    #
    else:
        print('O erro angular = {} está fora da tolerancia = {}, é indicado voltar a campo ou conferir os dados'.format(EaU,TaU)) 
        f1=input('Você deseja fazer o procedimento mesmo assim? (Y) para sim e (N) para não: ')
        if f1=='N':
            exit() #Comando para o programa se encerrar caso essa opção seja selecionada
    #
    #
    #
    print('------------------------------------------------------------------------------------------------------------------------------------------')

    #
    #
    #
    #
    #
    #Seção 2: Compensação do erro angular com variação de método, Método 1 = Classica divisão do erro entre todos os pontos | Método 2 = Divisão do erro nos ultimos pontos considerando o acumulo de erro

    aSAngHC=[]                      #aSAngHC = array salvar angulo horizontal corrigido
    aSAngHCU=[]                     #aSAngHCU = array salvar angulo horizontal corrigido para usuário
    up=[]                           #linha que guarda os angulos horizontais corridos no método 2 de correção
    while True:
        try:
            mcea=input('Métodos de compensação do erro angular:\nMétodo(1): Divisão igualitaria do erro entre os pontos\nMétodo(2): Divisão do erro entre os ultimos pontos considerando o acumulo de erro\nSelecione o método (1) ou (2): ')
            if not(mcea == '1' or mcea == '2'):
                raise ValueError('Selecione a opção (1) ou (2)')
            break
        except ValueError as error:
            print(error)
    #
    #
    #
    if mcea == '1':
        Ead=float(-Ea/Pn)                                                    #Ead = Erro angular distribuido
        for i in aSAngHDP:
            c=i+Ead;aSAngHC.append(c);                                      #somando a correção do erro nos angulos e salvando em aSAngHC
    #
    #
    #
    if mcea == '2':
        lp=float(0.3);cp=round((int(Pn*lp))+0.5);Ead=float(-Ea/cp);          #lp é porcentagem de pontos selecionados e cp é o número de pontos representados por esse valor, Ead é o erro dividido entre os pontos selecionados 
        ic = len(aSAngHDP)-cp;v=aSAngHDP[0:ic]                               #Para realizar essa correção é importante que apenas o ultimos valores dos angulos sejam corrigidos, portanto ic é a variavel que a guarda a posição de inicio em que os pontos serão corrigidos, v guarda todos os valores da linha até a posição gerada pela variavel ic, realizando um "corte" no vetor
        for i in aSAngHDP[ic:]:                                              #Esse for faz com que somente os valores a frente da posição ic sejam corrigidos, esses valores são salvos em up
            c=i+Ead;up.append(c);       
        aSAngHC=(v+up)                                                      #aqui nós pegamos o vetor cortado e juntamos com a parte corrigida que alteramos
    #
    #
    #
    for i in range(len(aSAngHC)):                                           #Este for lê a quantidade de itens na linha e mede seu alcance com o comando range, logo esse for denomina que o ciclo de repetição se estenderá até todos os itens da lista 
        angHCU=llp.to_deg_min_sec(aSAngHC[i]);angHCU='{}\u00B0{}\u02BC{}\u02EE'.format(int(angHCU[0]),int(angHCU[1]),float(angHCU[2]));aSAngHCU.append(angHCU)
    print(aSAngHCU)
    print('------------------------------------------------------------------------------------------------------------------------------------------')

    #
    #
    #
    #
    #

    #Seção 2.1 calculo dos azimutes
    while True:
        try: 
            AzPP  = input('Insira o azimute de partida: ');AzPP=AzPP.split("\u00B0");AzPP1=(int(AzPP[0]),int(AzPP[1]),float(AzPP[2]));AzPP=llp.to_dec_deg(AzPP1[0],AzPP1[1],AzPP1[2]);AzPPU=llp.to_deg_min_sec(AzPP)
            AzPPU='{}\u00B0{}\u02BC{}\u02EE'.format(int(AzPPU[0]),int(AzPPU[1]),float(AzPPU[2]))
            if not(AzPP1[0] < n[0] and AzPP1[1] < n[1] and AzPP1[2] < n[2]):                  #Este if cria a condição de escolha para o programa identificar que o valor inserado está errado
                raise ValueError('Valor fora do intervalo permitido')
            break
        except ValueError as error:
            print(error)
    AzPPi = AzPP                          #Guardando o azimute inicial em outra variavel para manipular o laço de repetição
    aSAz=[]                               #array que guarda os azimutes calculados
    aSAngHDP=[float(x) for x in aSAngHDP]
    aSAngHDPaz=aSAngHDP[1:]
    aSAzU=[]
    #
    #
    #
    for i in range(len(aSAngHDPaz)):
        AzPPi=AzPPi+aSAngHDPaz[i]-180;
        if AzPPi >= 360:
            AzPPi=AzPPi-360
            aSAz.append(AzPPi);
        if AzPPi < 0:
            AzPPi=AzPPi+360
            aSAz.append(AzPPi);
        aSAz.append(AzPPi);
    azf=(float(aSAz[-1]))+(float(aSAngHDP[0]))-180
    #
    #
    #
    if azf >= 360:
        azf=azf-360
    #
    #
    #
    if azf < 0:
        azf=azf+360
    #
    #
    #
    azf=round(azf,5)
    if azf==AzPP:
        print('Azimutes calculados corretamente')
        aSAz=[azf]+aSAz
    #
    #
    #
    else: print('O azimute inicial calculado = {}, não bate com o inserido pelo usuário'.format(azf));aSAz=[azf]+aSAz;
    #
    #
    #
    for i in range(len(aSAz)):
        aSAzUi=llp.to_deg_min_sec(aSAz[i]);
        aSAzUia=round(aSAzUi[2],3);aSAzUic=llp.to_dec_deg(aSAzUi[0],aSAzUi[1],aSAzUia);aSAzUi=llp.to_deg_min_sec(aSAzUic)
        aSAzUi='{}\u00B0{}\u02BC{}\u02EE'.format(int(aSAzUi[0]),int(aSAzUi[1]),float(aSAzUi[2]));aSAzU.append(aSAzUi)    
    print('azimutes:{}'.format(aSAzU))
    print('------------------------------------------------------------------------------------------------------------------------------------------')

    #
    #
    #
    #
    #
    #Seção 2.2 calculo das coordenadas provisórias e finais
    while True:
        try:
            xPP = float(input('Insira coordenada conhecida do ponto X em metros: '));yPP = float(input('Insira coordenada conhecida do ponto Y em metros: '))
            if not(type(xPP) and type(yPP) is float):
                raise ValueError('insira o valor corretamente')
            break
        except ValueError as error:
            print(error)
    #
    #
    #
    Dhs=[] #linha que ficaram todas as distancias horizontais
    for i in range(Pn):
        while True:
            try:
                Dh = float(input('Insira a distancia horizontal do ponto em metros, em ordem de pontos: '));Dh=round(Dh,5);Dhs.append(Dh)
                if not(type(Dh) is float):
                    raise ValueError('Insira a distancia horizontal corretamente')
                break
            except ValueError as error:
                print(error)
    #
    #
    #
    while True: #while para inserção da tolerancia
        try:
            Tl = input('Insira tolerancia linear: ');Tl1=Tl.split('/');Tl=(int(Tl1[0])/int(Tl1[1]));Tl=float(Tl);
            if not(type(Tl) is float):
                raise ValueError('insira o valor corretamente')
            break
        except ValueError as error:
            print(error)
    #
    #
    #
    #calculo das coordenadas parciais
    xPs=[];yPs=[]
    xPPi=round((float(xPP)),5);yPPi=round((float(yPP)),5)    #Guarda valor inicial para coordenadas parciais
    xPPci=round((float(xPP)),5);yPPci=round((float(yPP)),5)  #Guarda valor inicial para coordenadas corrigidas
    for i in range(Pn-1):
        xPPi=xPPi+(Dhs[i]*m.sin(m.radians(aSAz[i])));xPPi=round(xPPi,5);xPs.append(xPPi)
        yPPi=yPPi+(Dhs[i]*m.cos(m.radians(aSAz[i])));yPPi=round(yPPi,5);yPs.append(yPPi)
    #
    #
    #
    #Calculo da coordenada conhecida calculada
    xPPc=xPs[-1]+(Dhs[-1]*m.sin(m.radians(aSAz[-1])));xPPc=round(xPPc,5)
    yPPc=yPs[-1]+(Dhs[-1]*m.cos(m.radians(aSAz[-1])));yPPc=round(yPPc,5)
    xPs=[xPPc]+xPs;yPs=[yPPc]+yPs
    #
    #
    #
    #erro linear
    ex=float(xPs[0]-xPP);ey=float(yPs[0]-yPP);ep=m.sqrt((ex**2)+(ey**2));ep=round(ep,5)
    z=float(sum(Dhs)/ep);
    if 1/z < Tl:
        print('O erro linear está abaixo da tolerancia')
    #
    #
    #
    #calculo das coordenadas corrigidas
    CxPs=[]
    CyPs=[]
    for i in range(Pn-1):
        xPPci=xPPci+(Dhs[i]*m.sin(m.radians(aSAz[i]))+((-ex*(Dhs[i]))/sum(Dhs)));xPPci=round(xPPci,5);CxPs.append(xPPci)
        yPPci=yPPci+(Dhs[i]*m.cos(m.radians(aSAz[i]))+((-ey*(Dhs[i]))/sum(Dhs)));yPPci=round(yPPci,5);CyPs.append(yPPci)
    CxPs=[xPP]+CxPs;CyPs=[yPP]+CyPs #linhas com resultado final das coordenadas corrigidas
    #
    #
    #
    os.chdir('C:/Users/Dio/Documents/visual code/Projetos') #comando para entrar na pasta escolhida
    os.chdir(dir)
    #
    #
    #        
    pnex=[]
    for i in range(Pn): #for para gerar a ordenação de pontos
        pnx=('P'+str(i));pnex.append(pnx)
    #
    #
    #
    # comando para ordenar os dados em um dataframe e associa-lo em uma variavel aplicar funcionalidades
    df = DataFrame({'Pontos': pnex, 'Angulos Horizontais': aSAngHCU,'Angulo Zenitais': aSAzU,'Distâncias': Dhs,'Coordenadas em X': CxPs,'Coordenadas em Y': CyPs})
    # comando para salvar um arquivo excel usando a funcionalidade do dataframe da biblioteca pandas
    df.to_excel('Projeto'+dir+'.xlsx', sheet_name='sheet1', index=False)
    #
    #
    #
    #Plot da imagem da poligonal
    plt.grid(True)  #Plot sai com as linhas de marcação
    book = xlrd.open_workbook(diretório) #Comando para ler o arquivo excel
    sh = book.sheet_by_index(0)        
    py=[];px=[];pp=[]
    Pn1=Pn+1
    for i in range(Pn1):   #for para extrair os dados do excel
        x=sh.cell_value(rowx=i, colx=4);px.append(x)
        y=sh.cell_value(rowx=i, colx=5);py.append(y)
        p=sh.cell_value(rowx=i, colx=0);pp.append(p)
    pp=pp[1:]   #todos os valores começam a partir do [1] pois o item [0] é o nome da lista
    px=px[1:]
    py=py[1:]
    #
    #
    #
    #Retas
    plotp=plt.plot(px,py,':c') #Os ":" fazem as linhas ficarem tracejadas, enquanto a letra "c" representa a cor Ciano
    for i in range(len(px)):    #For para inserir o texto correspondente a cada ponto
        plt.text(px[i], py[i], pp[i])
    #
    #
    #
    plotp=plt.plot([px[0],px[-1]],[py[0],py[-1]],':c') #Este plot está relacionado a ligação do ultimo ponto com o primeiro
    plt.scatter(px,py,color='red')          #comando para colocar pontos de sinalização em cada vertice
    plt.title('Poligonal '+dir)             #titulo da imagem
    plt.xlabel('Coordenada X em metros')    #nomeando os eixos
    plt.ylabel('Coordenada Y em metros')
    #
    #
    #
    plt.savefig('Poligonal'+dir)            #salvar o arquivo em PNG
    plt.show()                              #comando final após realizar todos as funções na tabela, exibindo uma só imagem com as implementações
    exit()

start();crpj(criarprojeto);select();poligonal(sel)   #O código posssui funções(def) que necessitam ser chamadas para serem executadas, em algumas funções está entre parenteses a variavel que se quer conservar durante o a função pro restante do código