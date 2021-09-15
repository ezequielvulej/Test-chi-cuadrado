# -*- coding: utf-8 -*-
"""
Created on Sun May 24 16:59:34 2020

@author: evule
"""

import pandas as pd
import xlrd
import numpy as np
import scipy.stats as sc
import math

Libro=xlrd.open_workbook('Ajuste distribucion.xlsx')
Inputs=Libro.sheet_by_name('Inputs')

FechaInicio=xlrd.xldate_as_datetime(Inputs.cell(0,1).value, datemode=0)
FechaFin=xlrd.xldate_as_datetime(Inputs.cell(1,1).value, datemode=0)

Serie=pd.read_excel('Ajuste distribucion.xlsx', sheet_name='Serie')
Serie=Serie[(Serie['Fecha']>=FechaInicio) & (Serie['Fecha']<=FechaFin)]


n=len(Serie)
k=int(round(1+np.log10(n)/np.log10(2),0)) #Defino intervalos utilizando la regla de Sturges
Minimo=min(Serie.Valores)
Maximo=max(Serie.Valores)
Amplitud=(Maximo-Minimo)/k

IntervaloMax=pd.Series([(Minimo+Amplitud*k) for k in range(1,k+1)])
IntervaloMax.reset_index(drop=True,inplace=True)

IntervaloMin=pd.Series(Minimo).append(IntervaloMax.iloc[:-1])
IntervaloMin.reset_index(drop=True,inplace=True)

Intervalo=pd.concat([IntervaloMin,IntervaloMax], axis=1)
Intervalo.columns=['Minimo','Maximo']

def Cuenta(y,z): #Para contar la cantidad de datos real en cada intervalo
    return (sum((Serie.Valores>=y)&(Serie.Valores<=z)))

ni=pd.Series([Cuenta(Intervalo.Minimo[i],Intervalo.Maximo[i]) for i in range(0,k)])

Cortar=sum(np.cumsum(ni.iloc[::-1])<=5)-1 #Cuando el ultimo intervalo tiene menos de 5 datos, los agrupo
Intervalo=Intervalo.iloc[:-Cortar]
Intervalo.iloc[-1]['Maximo']=Maximo

ni=pd.Series([Cuenta(Intervalo.Minimo[i],Intervalo.Maximo[i]) for i in range(0,len(Intervalo))])
Intervalo=pd.concat([Intervalo,ni], axis=1)
Intervalo.columns=['Minimo','Maximo','ni']

Intervalo['Xi']=(Intervalo.Minimo+Intervalo.Maximo)/2 #Punto medio del intervalo
Intervalo['Frec relativa']=Intervalo['ni']/n

nint=len(Intervalo)
Alpha=Inputs.cell(2,1).value

#Test chi-cuadrado para distintas distribuciones:
MuDistrNormal=Serie.Valores.mean()
SigmaDistrNormal=Serie.Valores.std()
Intervalo['ProbEsperadaDistrNormal']=sc.norm(MuDistrNormal,SigmaDistrNormal).cdf(Intervalo.Maximo)-sc.norm(MuDistrNormal,SigmaDistrNormal).cdf(Intervalo.Minimo)
Intervalo.loc[0,'ProbEsperadaDistrNormal']=sc.norm(MuDistrNormal,SigmaDistrNormal).cdf(Intervalo.loc[0,'Maximo'])
Intervalo.loc[Intervalo.index[-1],'ProbEsperadaDistrNormal']=1-sum(Intervalo['ProbEsperadaDistrNormal'][:-1]) #Ajusto por ley de cierre
Intervalo['FrecEsperadaDistrNormal']=Intervalo['ProbEsperadaDistrNormal']*n
Intervalo['AuxCostoDistrNormal']=((Intervalo['ni']-Intervalo['FrecEsperadaDistrNormal'])**2)/Intervalo['FrecEsperadaDistrNormal']
CantParamDistrNormal=2
GLDistrNormal=nint-CantParamDistrNormal-1 #Grados de libertad
CriticoDistrNormal=sc.chi2.ppf(1-Alpha,GLDistrNormal)
ResultadoDistrNormal=sum(Intervalo['AuxCostoDistrNormal'])<CriticoDistrNormal


LambdaDistrGamma=Serie.Valores.mean()/(Serie.Valores.std()**2)
KDistrGamma=Serie.Valores.mean()*LambdaDistrGamma
Intervalo['ProbEsperadaDistrGamma']=sc.gamma(scale=1/LambdaDistrGamma, a=KDistrGamma).cdf(Intervalo.Maximo)-sc.gamma(scale=1/LambdaDistrGamma, a=KDistrGamma).cdf(Intervalo.Minimo)
Intervalo.loc[0,'ProbEsperadaDistrGamma']=sc.gamma(scale=1/LambdaDistrGamma, a=KDistrGamma).cdf(Intervalo.loc[0,'Maximo'])
Intervalo.loc[Intervalo.index[-1],'ProbEsperadaDistrGamma']=1-sum(Intervalo['ProbEsperadaDistrGamma'][:-1])
Intervalo['FrecEsperadaDistrGamma']=Intervalo['ProbEsperadaDistrGamma']*n
Intervalo['AuxCostoDistrGamma']=((Intervalo['ni']-Intervalo['FrecEsperadaDistrGamma'])**2)/Intervalo['FrecEsperadaDistrGamma']
CantParamDistrGamma=2
GLDistrGamma=nint-CantParamDistrGamma-1
CriticoDistrGamma=sc.chi2.ppf(1-Alpha,GLDistrGamma)
ResultadoDistrGamma=sum(Intervalo['AuxCostoDistrGamma'])<CriticoDistrGamma



SigmaDistrLognormal=(math.log(((Serie.Valores)).mean()**2+(Serie.Valores.std())**2)-2*math.log(Serie.Valores.mean()))**0.5
MuDistrLognormal=math.log(Serie.Valores.mean())-SigmaDistrLognormal**2/2
Intervalo['ProbEsperadaDistrLognormal']=sc.lognorm(scale=np.exp(MuDistrLognormal),s=SigmaDistrLognormal).cdf(Intervalo.Maximo)-sc.lognorm(scale=np.exp(MuDistrLognormal),s=SigmaDistrLognormal).cdf(Intervalo.Minimo)
Intervalo.loc[0,'ProbEsperadaDistrLognormal']=sc.lognorm(scale=np.exp(MuDistrLognormal),s=SigmaDistrLognormal).cdf(Intervalo.loc[0,'Maximo'])
Intervalo.loc[Intervalo.index[-1],'ProbEsperadaDistrLognormal']=1-sum(Intervalo['ProbEsperadaDistrLognormal'][:-1])
Intervalo['FrecEsperadaDistrLognormal']=Intervalo['ProbEsperadaDistrLognormal']*n
Intervalo['AuxCostoDistrLognormal']=((Intervalo['ni']-Intervalo['FrecEsperadaDistrLognormal'])**2)/Intervalo['FrecEsperadaDistrLognormal']
CantParamDistrLognormal=2
GLDistrLognormal=nint-CantParamDistrLognormal-1
CriticoDistrLognormal=sc.chi2.ppf(1-Alpha,GLDistrLognormal)
ResultadoDistrLognormal=sum(Intervalo['AuxCostoDistrLognormal'])<CriticoDistrLognormal





BetaDistrGumbel=Serie.Valores.std()*math.sqrt(6)/math.pi
MuDistrGumbel=Serie.Valores.mean()-BetaDistrGumbel*0.577215664901532
Intervalo['ProbEsperadaDistrGumbel']=sc.gumbel_r(MuDistrGumbel,BetaDistrGumbel).cdf(Intervalo.Maximo)-sc.gumbel_r(MuDistrGumbel,BetaDistrGumbel).cdf(Intervalo.Minimo)
Intervalo.loc[0,'ProbEsperadaDistrGumbel']=sc.gumbel_r(MuDistrGumbel,BetaDistrGumbel).cdf(Intervalo.loc[0,'Maximo'])
Intervalo.loc[Intervalo.index[-1],'ProbEsperadaDistrGumbel']=1-sum(Intervalo['ProbEsperadaDistrGumbel'][:-1])
Intervalo['FrecEsperadaDistrGumbel']=Intervalo['ProbEsperadaDistrGumbel']*n
Intervalo['AuxCostoDistrGumbel']=((Intervalo['ni']-Intervalo['FrecEsperadaDistrGumbel'])**2)/Intervalo['FrecEsperadaDistrGumbel']
CantParamDistrGumbel=2
GLDistrGumbel=nint-CantParamDistrGumbel-1
CriticoDistrGumbel=sc.chi2.ppf(1-Alpha,GLDistrGumbel)
ResultadoDistrGumbel=sum(Intervalo['AuxCostoDistrGumbel'])<CriticoDistrGumbel



MuDistrLogistica=Serie.Valores.mean()
SDistrLogistica=Serie.Valores.std()*(3**0.5)/math.pi
Intervalo['ProbEsperadaDistrLogistica']=sc.logistic(MuDistrLogistica,SDistrLogistica).cdf(Intervalo.Maximo)-sc.logistic(MuDistrLogistica,SDistrLogistica).cdf(Intervalo.Minimo)
Intervalo.loc[0,'ProbEsperadaDistrLogistica']=sc.logistic(MuDistrLogistica,SDistrLogistica).cdf(Intervalo.loc[0,'Maximo'])
Intervalo.loc[Intervalo.index[-1],'ProbEsperadaDistrLogistica']=1-sum(Intervalo['ProbEsperadaDistrLogistica'][:-1])
Intervalo['FrecEsperadaDistrLogistica']=Intervalo['ProbEsperadaDistrLogistica']*n
Intervalo['AuxCostoDistrLogistica']=((Intervalo['ni']-Intervalo['FrecEsperadaDistrLogistica'])**2)/Intervalo['FrecEsperadaDistrLogistica']
CantParamDistrLogistica=2
GLDistrLogistica=nint-CantParamDistrLogistica-1
CriticoDistrLogistica=sc.chi2.ppf(1-Alpha,GLDistrLogistica)
ResultadoDistrLogistica=sum(Intervalo['AuxCostoDistrLogistica'])<CriticoDistrLogistica

AlphaRes=0.01
#La hipotesis nula nunca se acepta. El valor entre comillas (se "acepta") en realidad significa que no se rechaza.
Resumen=pd.DataFrame({'Distribucion':['Normal','Gamma','Lognormal','Gumbel'], 'Se "acepta" el test?':[ResultadoDistrNormal,ResultadoDistrGamma,ResultadoDistrLognormal,ResultadoDistrGumbel]}, columns=['Distribucion','Se "acepta" el test?','Cota inferior 1%', 'Cota superior 1%'])
Resumen.loc[0,'Cota inferior 1%']=sc.norm(MuDistrNormal,SigmaDistrNormal).ppf(AlphaRes)
Resumen.loc[0,'Cota superior 1%']=sc.norm(MuDistrNormal,SigmaDistrNormal).ppf(1-AlphaRes)
Resumen.loc[1,'Cota inferior 1%']=sc.gamma(scale=1/LambdaDistrGamma, a=KDistrGamma).ppf(AlphaRes)
Resumen.loc[1,'Cota superior 1%']=sc.gamma(scale=1/LambdaDistrGamma, a=KDistrGamma).ppf(1-AlphaRes)
Resumen.loc[2,'Cota inferior 1%']=sc.lognorm(scale=np.exp(MuDistrLognormal),s=SigmaDistrLognormal).ppf(AlphaRes)
Resumen.loc[2,'Cota superior 1%']=sc.lognorm(scale=np.exp(MuDistrLognormal),s=SigmaDistrLognormal).ppf(1-AlphaRes)
Resumen.loc[3,'Cota inferior 1%']=sc.gumbel_r(MuDistrGumbel,BetaDistrGumbel).ppf(AlphaRes)
Resumen.loc[3,'Cota superior 1%']=sc.gumbel_r(MuDistrGumbel,BetaDistrGumbel).ppf(1-AlphaRes)
print(Resumen)

