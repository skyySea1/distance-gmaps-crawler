from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options
import openpyxl
import time

service = EdgeService()

options = webdriver.EdgeOptions()
# options.add_argument('headless')  

       
driver = webdriver.Edge( service=service, options=options)

d_metropolitana =  """CAMAÇARI, Bahia
CANDEIAS, Bahia
DIAS d'ÁVILA, Bahia
ITAPARICA, Bahia
LAURO DE FREITAS, Bahia
MADRE DE DEUS, Bahia
MATA DE SÃO JOÃO, Bahia
POJUCA, Bahia
SALVADOR, Bahia
SÃO FRANCISCO DO CONDE, Bahia
SÃO SEBASTIÃO DO PASSÉ, Bahia
SIMÕES FILHO, Bahia
VERA CRUZ, Bahia
"""
d_centronorte = """ ABAÍRA, Bahia
ÁGUA FRIA, Bahia
AMARGOSA, Bahia
AMÉLIA RODRIGUES, Bahia
ANDARAÍ, Bahia
ANGUERA, Bahia
ANTÔNIO CARDOSO, Bahia
BAIXA GRANDE, Bahia
BARRA DA ESTIVA, Bahia
BOA VISTA DO TUPIM, Bahia
BONINAL, Bahia
BONITO, Bahia
BREJÕES, Bahia
CABACEIRAS DO PARAGUAÇU, Bahia
CACHOEIRA, Bahia
CAPELA DO ALTO ALEGRE, Bahia
CAPIM GROSSO, Bahia
CASTRO ALVES, Bahia
CONCEIÇÃO DA FEIRA, Bahia
CONCEIÇÃO DO ALMEIDA, Bahia
CONCEIÇÃO DO JACUÍPE, Bahia
CORAÇÃO DE MARIA, Bahia
CRAVOLÂNDIA, Bahia
CRUZ DAS ALMAS, Bahia
DOM MACEDO COSTA, Bahia
ELÍSIO MEDRADO, Bahia
FEIRA DE SANTANA, Bahia
GAVIÃO, Bahia
GOVERNADOR MANGABEIRA, Bahia
IAÇU, Bahia
IBICOARA, Bahia
IBIQUERA, Bahia
IBITIARA, Bahia
IPECAETÁ, Bahia
IPIRÁ, Bahia
IRAJUBA, Bahia
IRAMAIA, Bahia
IRAQUARA, Bahia
IRARÁ, Bahia
ITABERABA, Bahia
ITAETÉ, Bahia
ITAQUARA, Bahia
ITATIM, Bahia
ITIRUÇU, Bahia
JAGUAQUARA, Bahia
JIQUIRIÇÁ, Bahia
JUSSIAPE, Bahia
LAFAIETE COUTINHO, Bahia
LAJE, Bahia
LAJEDINHO, Bahia
LAJEDO DO TABOCAL, Bahia
LENÇÓIS, Bahia
MACAJUBA, Bahia
MAIRI, Bahia
MARACÁS, Bahia
MARAGOGIPE, Bahia
MARCIONÍLIO SOUZA, Bahia
MILAGRES, Bahia
MORRO DO CHAPÉU, Bahia
MUCUGÊ, Bahia
MUNDO NOVO, Bahia
MUNIZ FERREIRA, Bahia
MURITIBA, Bahia
MUTUÍPE, Bahia
NAZARÉ, Bahia
NOVA FÁTIMA, Bahia
NOVA ITARANA, Bahia
NOVA REDENÇÃO, Bahia
NOVO HORIZONTE, Bahia
PALMEIRAS, Bahia
PÉ DE SERRA, Bahia
PIATÃ, Bahia
PINTADAS, Bahia
PIRITIBA, Bahia
PLANALTINO, Bahia
QUIXABEIRA, Bahia
RAFAEL JAMBEIRO, Bahia
RIACHÃO DO JACUÍPE, Bahia
RIO DE CONTAS, Bahia
SALINAS DA MARGARIDA, Bahia
SANTA BÁRBARA, Bahia
SANTA INÊS, Bahia
Santa Terezinha, Bahia
SANTANÓPOLIS, Bahia
SANTO AMARO, Bahia
SANTO ANTÔNIO DE JESUS, Bahia
SANTO ESTÊVÃO, Bahia
SÃO FELIPE, Bahia
SÃO FÉLIX, Bahia
SÃO GONÇALO DOS CAMPOS, Bahia
SÃO JOSÉ DO JACUÍPE, Bahia
SÃO MIGUEL DAS MATAS, Bahia
SAPEAÇU, Bahia
SAUBARA, Bahia
SEABRA, Bahia
SERRA PRETA, Bahia
SOUTO SOARES, Bahia
TANQUINHO, Bahia
TAPIRAMUTÁ, Bahia
TEODORO SAMPAIO, Bahia
TERRA NOVA, Bahia
UBAÍRA, Bahia
UTINGA, Bahia
VÁRZEA DA ROÇA, Bahia
VÁRZEA DO POÇO, Bahia
VARZEDO, Bahia
WAGNER, Bahia """
d_centrosul =""" AIQUARA, Bahia
ANAGÉ, Bahia
APUAREMA, Bahia
ARACATU, Bahia
BARRA DO CHOÇA, Bahia
BARRA DO ROCHA, Bahia
BELO CAMPO, Bahia
BOA NOVA, Bahia
BOM JESUS DA SERRA, Bahia
BRUMADO, Bahia
CAATIBA, Bahia
CACULÉ, Bahia
CAETANOS, Bahia
CAETITÉ, Bahia
CANDIBA, Bahia
CÂNDIDO SALES, Bahia
CARAÍBAS, Bahia
CONDEÚBA, Bahia
CONTENDAS DO SINCORÁ, Bahia
CORDEIROS, Bahia
DÁRIO MEIRA, Bahia
DOM BASÍLIO, Bahia
ENCRUZILHADA, Bahia
FIRMINO ALVES, Bahia
GONGOGI, Bahia
GUAJERU, Bahia
GUANAMBI, Bahia
IBIASSUCÊ, Bahia
IBICUÍ, Bahia
IBIRATAIA, Bahia
IGUAÍ, Bahia
IPIAÚ, Bahia
ITAGI, Bahia
ITAGIBÁ, Bahia
ITAMARI, Bahia
ITAMBÉ, Bahia
ITAPETINGA, Bahia
ITARANTIM, Bahia
ITORORÓ, Bahia
ITUAÇU, Bahia
IUIU, Bahia
JACARACI, Bahia
JEQUIÉ, Bahia
JITAÚNA, Bahia
LAGOA REAL, Bahia
LICÍNIO DE ALMEIDA, Bahia
LIVRAMENTO DE NOSSA SENHORA, Bahia
MACARANI, Bahia
MAETINGA, Bahia
MAIQUINIQUE, Bahia
MALHADA DE PEDRAS, Bahia
MANOEL VITORINO, Bahia
MIRANTE, Bahia
MORTUGABA, Bahia
NOVA CANAÃ, Bahia
NOVA IBIÁ, Bahia
PALMAS DE MONTE ALTO, Bahia
PINDAÍ, Bahia
PIRIPÁ, Bahia
PLANALTO, Bahia
POÇÕES, Bahia
POTIRAGUÁ, Bahia
PRESIDENTE JÂNIO QUADROS, Bahia
RIBEIRÃO DO LARGO, Bahia
RIO DO ANTÔNIO, Bahia
SANTA CRUZ DA VITÓRIA, Bahia
SEBASTIÃO LARANJEIRAS, Bahia
TANHAÇU, Bahia
TANQUE NOVO, Bahia
TREMEDAL, Bahia
UBATÃ, Bahia
URANDI, Bahia
VITÓRIA DA CONQUISTA, Bahia"""
d_oeste = """ANGICAL, Bahia
BAIANÓPOLIS, Bahia
BARRA, Bahia
BARREIRAS, Bahia
BOM JESUS DA LAPA, Bahia
BOQUIRA, Bahia
BOTUPORÃ, Bahia
BREJOLÂNDIA, Bahia
BROTAS DE MACAÚBAS, Bahia
BURITIRAMA, Bahia
CANÁPOLIS, Bahia
CARINHANHA, Bahia
CATOLÂNDIA, Bahia
CATURAMA, Bahia
COCOS, Bahia
CORIBE, Bahia
CORRENTINA, Bahia
COTEGIPE, Bahia
CRISTÓPOLIS, Bahia
ÉRICO CARDOSO, Bahia
FEIRA DA MATA, Bahia
FORMOSA DO RIO PRETO, Bahia
IBIPITANGA, Bahia
IBOTIRAMA, Bahia
IGAPORÃ, Bahia
JABORANDI, Bahia
LUÍS EDUARDO MAGALHÃES, Bahia
MACAÚBAS, Bahia
MALHADA, Bahia
MANSIDÃO, Bahia
MATINA, Bahia
MORPARÁ, Bahia
MUQUÉM DO SÃO FRANCISCO, Bahia
OLIVEIRA DOS BREJINHOS, Bahia
PARAMIRIM, Bahia
PARATINGA, Bahia
RIACHÃO DAS NEVES, Bahia
RIACHO DE SANTANA, Bahia
RIO DO PIRES, Bahia
SANTA MARIA DA VITÓRIA, Bahia
SANTA RITA DE CÁSSIA, Bahia
SANTANA, Bahia
SÃO DESIDÉRIO, Bahia
SÃO FÉLIX DO CORIBE, Bahia
SERRA DO RAMALHO, Bahia
SERRA DOURADA, Bahia
SÍTIO DO MATO, Bahia
TABOCAS DO BREJO VELHO, Bahia
WANDERLEY, Bahia"""
d_norte = """AMÉRICA DOURADA, Bahia
ANDORINHA, Bahia
ANTÔNIO GONÇALVES, Bahia
BARRA DO MENDES, Bahia
BARRO ALTO, Bahia
CAÉM, Bahia
CAFARNAUM, Bahia
CALDEIRÃO GRANDE, Bahia
CAMPO ALEGRE DE LOURDES, Bahia
CAMPO FORMOSO, Bahia
CANARANA, Bahia
CANUDOS, Bahia
CASA NOVA, Bahia
CENTRAL, Bahia
CURAÇÁ, Bahia
FILADÉLFIA, Bahia
GENTIO DO OURO, Bahia
IBIPEBA, Bahia
IBITITÁ, Bahia
IPUPIARA, Bahia
IRECÊ, Bahia
ITAGUAÇU DA BAHIA, Bahia
JACOBINA, Bahia
JAGUARARI, Bahia
JOÃO DOURADO, Bahia
JUAZEIRO, Bahia
JUSSARA, Bahia
LAPÃO, Bahia
MIGUEL CALMON, Bahia
MIRANGABA, Bahia
MULUNGU DO MORRO, Bahia
OUROLÂNDIA, Bahia
PILÃO ARCADO, Bahia
PINDOBAÇU, Bahia
PONTO NOVO, Bahia
PRESIDENTE DUTRA, Bahia
REMANSO, Bahia
SÃO GABRIEL, Bahia
SAÚDE, Bahia
SENHOR DO BONFIM, Bahia
SENTO SÉ, Bahia
SERROLÂNDIA, Bahia
SOBRADINHO, Bahia
UAUÁ, Bahia
UIBAÍ, Bahia
UMBURANAS, Bahia
VÁRZEA NOVA, Bahia
XIQUE-XIQUE, Bahia"""
d_nordeste = """ABARÉ, Bahia
ACAJUTIBA, Bahia
ADUSTINA, Bahia
ALAGOINHAS, Bahia
ANTAS, Bahia
APORÁ, Bahia
ARAÇÁS, Bahia
ARACI, Bahia
ARAMARI, Bahia
BANZAÊ, Bahia
BARROCAS, Bahia
BIRITINGA, Bahia
CANDEAL, Bahia
CANSANÇÃO, Bahia
CARDEAL DA SILVA, Bahia
CATU, Bahia
CHORROCHÓ, Bahia
CÍCERO DANTAS, Bahia
CIPÓ, Bahia
CONCEIÇÃO DO COITÉ, Bahia
CONDE, Bahia
CORONEL JOÃO SÁ, Bahia
CRISÓPOLIS, Bahia
ENTRE RIOS, Bahia
ESPLANADA, Bahia
EUCLIDES DA CUNHA, Bahia
FÁTIMA, Bahia
GLÓRIA, Bahia
HELIÓPOLIS, Bahia
ICHU, Bahia
INHAMBUPE, Bahia
ITANAGRA, Bahia
ITAPICURU, Bahia
ITIÚBA, Bahia
JANDAÍRA, Bahia
JEREMOABO, Bahia
LAMARÃO, Bahia
MACURURÉ, Bahia
MONTE SANTO, Bahia
NORDESTINA, Bahia
NOVA SOURE, Bahia
NOVO TRIUNFO, Bahia
OLINDINA, Bahia
OURIÇANGAS, Bahia
PARIPIRANGA, Bahia
PAULO AFONSO, Bahia
PEDRÃO, Bahia
PEDRO ALEXANDRE, Bahia
QUEIMADAS, Bahia
QUIJINGUE, Bahia
RETIROLÂNDIA, Bahia
RIBEIRA DO AMPARO, Bahia
RIBEIRA DO POMBAL, Bahia
RIO REAL, Bahia
RODELAS, Bahia
SANTA BRÍGIDA, Bahia
SANTALUZ, Bahia
SÃO DOMINGOS, Bahia
SÁTIRO DIAS, Bahia
SERRINHA, Bahia
SÍTIO DO QUINTO, Bahia
TEOFILÂNDIA, Bahia
TUCANO, Bahia
VALENTE, Bahia"""
d_sul = """ALCOBAÇA, Bahia
ALMADINA, Bahia
Itabuna, Bahia
ARATUÍPE, Bahia
AURELINO LEAL, Bahia
Itabuna, Bahia
BELMONTE, Bahia
BUERAREMA, Bahia
Itabuna, Bahia
CAMACAN, Bahia
CAMAMU, Bahia
Itabuna, Bahia
CARAVELAS, Bahia
COARACI, Bahia
Itabuna, Bahia
FLORESTA AZUL, Bahia
GANDU, Bahia
Itabuna, Bahia
IBICARAÍ, Bahia
IBIRAPITANGA, Bahia
Itabuna, Bahia
IGRAPIÚNA, Bahia
ILHÉUS, Bahia
Itabuna, Bahia
ITABUNA, Bahia
ITACARÉ, Bahia
Itabuna, Bahia
ITAJU DO COLÔNIA, Bahia
ITAJUÍPE, Bahia
Itabuna, Bahia
ITANHÉM, Bahia
ITAPÉ, Bahia
Itabuna, Bahia
ITAPITANGA, Bahia
ITUBERÁ, Bahia
Itabuna, Bahia
JUCURUÇU, Bahia
JUSSARI, Bahia
Itabuna, Bahia
MARAÚ, Bahia
MASCOTE, Bahia
Itabuna, Bahia
MUCURI, Bahia
NILO PEÇANHA, Bahia
Itabuna, Bahia
PAU BRASIL, Bahia
PIRAÍ DO NORTE, Bahia
Itabuna, Bahia
PRADO, Bahia
PRESIDENTE TANCREDO NEVES, Bahia
Itabuna, Bahia
SANTA LUZIA, Bahia
SÃO JOSÉ DA VITÓRIA, Bahia
Itabuna, Bahia
TEIXEIRA DE FREITAS, Bahia
TEOLÂNDIA, Bahia
Itabuna, Bahia
UNA, Bahia
URUÇUCA, Bahia
Itabuna, Bahia
VEREDA, Bahia
WENCESLAU GUIMARÃES, Bahia"""

origem_cn = "Feira de Santana, Bahia"
origem_cs = "Vitória da Conquista, Bahia"
origem_o = "Barreiras, Bahia"
origem_n = "Juazeiro, Bahia"
origem_nd = "Serrinha, Bahia"
origem_s = "Itabuna, Bahia"
origem_m = "Salvador, Bahia"

# Definição das variáveis
origens = [origem_cn, origem_cs, origem_o, origem_n, origem_nd, origem_s, origem_m]
destinos =  d_centronorte.split("\n") +  d_centrosul.split("\n") + d_oeste.split("\n") + d_norte.split("\n") + d_nordeste.split("\n") + d_sul.split("\n") +  d_metropolitana.split("\n") 
origem = origens
# Loop para cada combinação de origem e destino
for origem in origens:
    for destino in destinos:
        # Abrir página inicial
        driver.get("https://www.google.com/maps/dir/Salvador,+Bahia/Abaíra,+BA")
        wait = WebDriverWait(driver, 15)  # Espera até 10 segundos
        element = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='omnibox-directions']/div/div[2]/div/div/div/div[2]/button")))
        element.click()


        origem_input = driver.find_element(By.XPATH, '//*[@id="sb_ifc50"]/input')
        origem_input.clear()
        origem_input.send_keys(origem)

        
        # Localizar e preencher o campo de destino
        destino_input = driver.find_element(By.XPATH, '//*[@id="sb_ifc51"]/input')
        destino_input.clear()
        destino_input.send_keys(destino)
        destino_input.send_keys(Keys.ENTER)
        
        distancia = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="section-directions-trip-0"]/div[1]/div/div[1]/div[2]/div')))

    
        
      
        
        # Aguardar tempo suficiente para o observer capturar as mudanças (ajuste conforme necessário)
        time.sleep(10)
        


# Fechar o navegador ao finalizar
driver.quit()
