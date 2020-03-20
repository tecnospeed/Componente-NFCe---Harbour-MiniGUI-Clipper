#include <hmg.ch>

Function Main

        Load Window Main
        Main.Center
        Main.Activate
	
Return

Function Main_BtnLoadConfig_Action
  
  NFCeX := CreateObject( "NFCeX.spdNFCeX" )
  
  //Configuração utilizando informações do arquivo .ini
  NFCeX:LoadConfig("nfceConfig.ini") //Informe o caminho do seu nfceConfig.ini

  //Também é possível configurar informando diretamente cada propriedade do componente, por ex:
  NFCeX:VersaoManual       := "6.0"
  NFCeX:TipoCertificado    := 2 //0=Memory, 1=LocalMachine, 2=File, 3=ActiveDirectory, 4=SmartCard
  NFCeX:CNPJ               := "" //Informe aqui o seu CNPJ
  NFCeX:NomeCertificado    := "" //Informe aqui o nome do seu certificado
  NFCeX:DiretorioTemplates := alltrim("C:\Program Files\TecnoSpeed\NFCe\Arquivos\Templates\vm60\")
  NFCeX:DiretorioEsquemas  := alltrim("C:\Program Files\TecnoSpeed\NFCe\Arquivos\Esquemas\vm60\")
  NFCeX:Ambiente           := 2 //1=Producao, 2=Homologacao

  Main.edtRetorno.value := "Configurações do componente carregadas ..."

  NFCeX := nil

Return Nil

Function Main_BtnStatusServico_Action

  NFCeX := CreateObject( "NFCeX.spdNFCeX" )
  NFCeX:LoadConfig("nfceConfig.ini") //Informe o caminho do seu nfceConfig.ini

  Main.edtRetorno.value := NFCeX:StatusDoServico()

  NFCeX := nil

Return Nil

Function Main_BtnGerarViaTx2_Action

  NFCeX := CreateObject( "NFCeX.spdNFCeX" )
  NFCeX:LoadConfig("nfceConfig.ini") //Informe o caminho do seu nfceConfig.ini

  Main.edtRetorno.value := NFCeX:ConverterLoteParaXml("Exemplo TX2 NFCe_Autorizando.txt", 1, "pl_009")

  NFCeX := nil

Return Nil

Function Main_BtnAssinarXml_Action

  if len(Main.edtRetorno.value) > 0
     NFCeX := CreateObject( "NFCeX.spdNFCeX" )
     NFCeX:LoadConfig("nfceConfig.ini") //Informe o caminho do seu nfceConfig.ini

     Main.edtRetorno.value := NFCeX:AssinarNota(Main.edtRetorno.value)

     NFCeX := nil
  else 
     Alert ("Informe antes um xml válido")
     Main.edtRetorno.SetFocus()
  endif

Return Nil

Function Main_BtnPreverDanfce_Action

  if len(Main.edtRetorno.value) > 0
     NFCeX := CreateObject( "NFCeX.spdNFCeX" )
     NFCeX:LoadConfig("nfceConfig.ini") //Informe o caminho do seu nfceConfig.ini

     NFCeX:PreverDanfce(Main.edtRetorno.value, "")
  
     NFCeX := nil
  else 
     Alert ("Informe antes um xml válido")
     Main.edtRetorno.SetFocus()
  endif

Return Nil

Function Main_BtnEnviarNFCe_Action

  if len(Main.edtRetorno.value) > 0
     NFCeX := CreateObject( "NFCeX.spdNFCeX" )
     NFCeX:LoadConfig("nfceConfig.ini") //Informe o caminho do seu nfceConfig.ini

     Main.edtRetorno.value := NFCeX:EnviarNFSincrono("1", Main.edtRetorno.value, 0)

     NFCeX := nil
  else 
     Alert ("Informe antes um xml válido")
     Main.edtRetorno.SetFocus()
  endif

Return Nil

Function Main_BtnConsultarNFCe_Action

  if len(Main.edtChaveNFCe.value) = 44
     NFCeX := CreateObject( "NFCeX.spdNFCeX" )
     NFCeX:LoadConfig("nfceConfig.ini") //Informe o caminho do seu nfceConfig.ini

     Main.edtRetorno.value := NFCeX:ConsultarNF(Main.edtChaveNFCe.value)

     NFCeX := nil
  else 
     Alert ("Informe antes uma chave válida com 44 digitos")
     Main.edtChaveNFCe.SetFocus()
  endif

Return Nil

Function Main_BtnExportarDanfce_Action

  if len(Main.edtRetorno.value) > 0
     NFCeX := CreateObject( "NFCeX.spdNFCeX" )
     NFCeX:LoadConfig("nfceConfig.ini") //Informe o caminho do seu nfceConfig.ini

     NFCeX:ExportarDanfce("1", Main.edtRetorno.value, "", 1, "")

     NFCeX := nil
  else 
     Alert ("Informe antes um xml autorizado válido")
     Main.edtRetorno.SetFocus()
  endif

Return Nil

Function Main_BtnListarCertificados_Action

  NFCeX := CreateObject( "NFCeX.spdNFCeX" )
  NFCeX:LoadConfig("nfceConfig.ini") //Informe o caminho do seu nfceConfig.ini

  Main.edtRetorno.value := NFCeX:ListarCertificados("|")

  NFCeX := nil

Return Nil

