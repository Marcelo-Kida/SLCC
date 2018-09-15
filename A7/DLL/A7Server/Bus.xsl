<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:LocalFunctions="urn:user-namespace-here" exclude-result-prefixes="msxsl LocalFunctions">
   <xsl:output method="xml" indent="no" />

   <msxsl:script language="VBScript" implements-prefix="LocalFunctions">
      <![CDATA[

                      Function flValueFromCSV(psCSV, psSeparador, plIndex, psTipo, pbObrigatorio)
                      Dim lvCSV
                      dim lsValor
                        
                        lvCSV = Split(psCSV, psSeparador)
                      
                        If UBound(lvCSV) < (plIndex - 1) Then
                            'Posição desejada não existe
                            flValueFromCSV = ""
                        Else
                      
                            lsValor = lvCSV(plIndex - 1)
                            
                            Select Case UCASE(psTipo)
                                  Case "STRING"
                                      
                                      lsValor = Trim(lsValor)
                                      
                                  Case "NUMBER"
                                  
                                      If pbObrigatorio Then
                                          If Replace(lsValor,"0","") = "" Then
                                              lsValor = ""
                                          End If  
                                      Else
                                          If Replace(lsValor,"0","") = "" Then
                                              lsValor = "0"
                                          End If  
                                      End If
                                      
                            End Select    
                            
                            flValueFromCSV = lsValor
                      
                        End If
                      
                      End Function
                      
                      Function flCsvToXML(psCSV, psSeparador)
                        
                        flCsvToXML  =  "<CSV><Item>" & Replace(psCSV,psSeparador,"</Item><Item>") & "</Item></CSV>"
                      
                      End Function
                        
                      Function flValorToSTR(psValor, psTipo, plTamanho, plDecimais, pbObrigatorio)
                      
                      Dim lsNumero 
                      Dim lsDecimal
                                              
                          Select Case Ucase(psTipo)
                            Case "STRING"
                            
                                flValorToSTR = Left(psValor & String(plTamanho, " "), plTamanho)
                                
                                If pbObrigatorio = "1" Then
                                  If Trim(flValorToSTR) = "" Then
                                      flValorToSTR = ""
                                  End If
                                else
                                  flValorToSTR = Left(psValor & String(plTamanho, " "), plTamanho)
                                End If
                                  
                            Case "NUMBER"
                            
                                If Not IsNumeric(plDecimais)  Then
                                   plDecimais = clng(0)
                                End If
                                
                                
                                If InStr(1, psValor, ",") > 0 Then
                                    
                                    lsNumero = Split(psValor, ",")(0)
                                    lsDecimal = Split(psValor, ",")(1)
                                     
                                    If abs(plDecimais) > 0 Then
                                    
                                        flValorToSTR = Right(String(plTamanho, "0") & Trim(lsNumero), plTamanho - plDecimais) & _
                                         Left(Trim(lsDecimal) & String(plDecimais, "0"), plDecimais)
                                     
                                    Else
                                    
                                        flValorToSTR = Right(String(plTamanho, "0") & Trim(lsNumero), plTamanho)
                                    
                                    End If
                                Else
                                    
                                    flValorToSTR = Right(String(plTamanho, "0") & Trim(psValor), plTamanho - plDecimais) & String(plDecimais, "0")
                                 
                                End If
                                
                                If pbObrigatorio = "1" Then
		                            If Trim(flValorToSTR) = "" or Replace(Replace(flValorToSTR,"0",""),",","") = "" Then
		                                flValorToSTR = ""
		                            End If
                                Else
	                                If InStr(1, psValor, ",") > 0 Then
	                                    
	                                    lsNumero = Split(psValor, ",")(0)
	                                    lsDecimal = Split(psValor, ",")(1)
	                                     
	                                    If abs(plDecimais) > 0 Then
	                                    
	                                        flValorToSTR = Right(String(plTamanho, "0") & Trim(lsNumero), plTamanho - plDecimais) & _
	                                         Left(Trim(lsDecimal) & String(plDecimais, "0"), plDecimais)
	                                     
	                                    Else
	                                    
	                                        flValorToSTR = Right(String(plTamanho, "0") & Trim(lsNumero), plTamanho)
	                                    
	                                    End If
	                                Else
	                                    
	                                    flValorToSTR = Right(String(plTamanho, "0") & Trim(psValor), plTamanho - plDecimais) & String(plDecimais, "0")
	                                 
	                                End If


                                End If

                          End Select
                          
                      End Function
                      
                      Function flValorToXML(psValor, psTipo, plTamanho, plDecimais, pbObrigatorio)

                      Dim lsNumero 
                      Dim lsDecimal
                      
                          Select Case Ucase(psTipo)
                              Case "STRING"
                              
                                  flValorToXML = Trim(psValor)
                              
                              Case "NUMBER"
                                  
                                  If Not IsNumeric(psValor)  Then
                                      flValorToXML =  psValor
                                      exit function   
                                  end if
                                  
                                  If Not IsNumeric(plDecimais)  Then
                                     plDecimais = clng(0)
                                  End If
                                                        
                                  If pbObrigatorio Then
                                      If Replace(psValor,"0","") = "" Then
                                          flValorToXML = "0"
                                      Else
                                        
                                          If plDecimais > 0 Then
                                              flValorToXML = FormatNumber(Left(psValor,plTamanho - plDecimais),0,,,0) & "," & Right(psValor, plDecimais)
                                          Else
                                              flValorToXML = (Left(psValor,plTamanho))
                                          End If
                                          
                                      End IF          
                                  Else
                                      If plDecimais > 0 Then
                                          flValorToXML = FormatNumber(Left(psValor,plTamanho - plDecimais),0,,,0) & "," & Right(psValor, plDecimais)
                                          
                                      Else
                                          flValorToXML = (Left(psValor,plTamanho))
                                      End If
                                  End IF          
                              
                          End Select

                      End Function
                      
                      Function flTrim(psString)
                          flTrim = Trim(psString)
                      End Function

                      Function flPreparaString(psString)
                          
                          If Trim(psString) = "" Then
                              flPreparaString = ""
                          Else
                              flPreparaString = psString
                          End If
                      
                      End Function
                      
                      function flValorXMLparaXML(psValor, psTipo, plDecimais, pbObrigatorio)
                          
                          if pbObrigatorio = "0" then  
                              
                              if len(trim(psValor)) > 0  then
                                  
                                  if asc(trim(psValor)) <> "10" then
                                      
                                      Select Case Ucase(psTipo)
                                          
                                          Case "STRING"
                                          
                                              flValorXMLparaXML = psValor
                                          
                                          Case "NUMBER"
                                                
                                              if plDecimais > 0 then
                                                  if instr(1,psValor,",") = 0 then  
                                              
                                                    flValorXMLparaXML = replace(trim(psValor),chr(10),"") & ",00" 
                                                    
                                                  else
                                                    flValorXMLparaXML = trim(psValor)
                                                  end if

                                              else
                                                  flValorXMLparaXML = psValor
                                              end if
                                          
                                      end select
                                      
                                      Exit function
                                  end if              
                              end if
                                
                              Select Case Ucase(psTipo)
                                  
                                  Case "STRING"
                                  
                                      flValorXMLparaXML = " "
                                  
                                  Case "NUMBER"
                                  
                                      if plDecimais > 0 then
                                          flValorXMLparaXML = "0,0"
                                      else
                                          flValorXMLparaXML = "0"
                                      end if
                                  
                              end select
                      
                          else
                            
                              Select Case Ucase(psTipo)
                                  
                                  Case "STRING"
                                      
                                      flValorXMLparaXML = psValor
                                  
                                  Case "NUMBER"
                                  
                                      if plDecimais > 0 then
                                          
                                          if instr(1,psValor,",") = 0 then  
                                      
                                            flValorXMLparaXML = replace(trim(psValor),chr(10),"") & ",00" 
                                            
                                          else
                                            flValorXMLparaXML = trim(psValor)
                                          end if
                                      else
                                          flValorXMLparaXML = psValor
                                      end if
                                  
                              end select
                                
                          end if    
                      
                      end function
                      

                ]]>
   </msxsl:script>

   <xsl:template match="/">
      <xsl:variable name="FormatoID" select="//Documento/Formato/IDOutPut" />

      <xsl:variable name="FormatoSTR" select="//Documento/Formato/STROutPut" />

      <xsl:variable name="FormatoXML" select="//Documento/Formato/XMLOutPut" />

      <xsl:variable name="FormatoCSV" select="//Documento/Formato/CSVOutPut" />

      <xsl:variable name="Mensagem" select="//Documento/Mensagem" />

      <xsl:variable name="OutPutXMLName" select="//Documento/Formato/XMLOutPut/@OutPutName" />

      <Saida xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="Evento.xsd">
<!--Identifica os Tipos de Entrada-->
         <xsl:choose>
<!-- 
**********************************************************************************************
Entrada do tipo String - Posicional                                                        
**********************************************************************************************
-->
            <xsl:when test="//Documento/Mensagem/@Tipo='String'">
<!-- Monta a Parte ID-->
               <xsl:if test="$FormatoID">
                  <xsl:element name="SaidaID">
                     <xsl:for-each select="$FormatoID/*">
                        <xsl:element name="{name()}">
                           <xsl:value-of select="LocalFunctions:flValorToSTR(substring($Mensagem, @Inicio, @TamanhoOriginal),string(@Tipo), number(@Tamanho), format-number(@Decimais,'0'), boolean(@Obrigatorio))" />
                        </xsl:element>
                     </xsl:for-each>
                  </xsl:element>
               </xsl:if>

<!-- Monta a Parte String-->
               <xsl:if test="$FormatoSTR">
                  <xsl:element name="SaidaSTR">
                     <xsl:for-each select="$FormatoSTR/*">
                        <xsl:element name="{name()}">
                           <xsl:value-of select="LocalFunctions:flValorToSTR(substring($Mensagem, @Inicio, @TamanhoOriginal),string(@Tipo), number(@Tamanho), format-number(@Decimais,'0'), boolean(@Obrigatorio))" />
                        </xsl:element>
                     </xsl:for-each>
                  </xsl:element>
               </xsl:if>

<!-- Monta a Parte XML-->
               <xsl:if test="$FormatoXML">
                  <xsl:element name="{$OutPutXMLName}">
                     <xsl:for-each select="$FormatoXML/*">
                        <xsl:element name="{name()}">
                           <xsl:value-of select="LocalFunctions:flValorToXML(substring($Mensagem, @Inicio, @Tamanho), string(@Tipo), number(@Tamanho), format-number(@Decimais,'0'), boolean(@Obrigatorio))" />
                        </xsl:element>
                     </xsl:for-each>
                  </xsl:element>
               </xsl:if>

<!-- Monta a Parte CSV-->
               <xsl:if test="$FormatoCSV">
                  <xsl:element name="SaidaCSV">
                     <xsl:for-each select="$FormatoCSV/*">
                        <xsl:element name="{name()}">
                           <xsl:value-of select="LocalFunctions:flValorToXML(substring($Mensagem, @Inicio, @Tamanho), string(@Tipo), number(@Tamanho), format-number(@Decimais,'0'), boolean(@Obrigatorio))" />
                        </xsl:element>

                        <xsl:if test="position() != last()">
                           <xsl:element name="Separador">
                              <xsl:value-of select="../@Delimitador" />
                           </xsl:element>
                        </xsl:if>
                     </xsl:for-each>
                  </xsl:element>
               </xsl:if>
            </xsl:when>

<!--
***********************************************************************************************
  Entrada do tipo XML
***********************************************************************************************
-->
            <xsl:when test="//Documento/Mensagem/@Tipo='XML'">
<!-- Monta a Parte ID-->
               <xsl:if test="$FormatoID">
                  <xsl:element name="SaidaID">
                     <xsl:for-each select="$FormatoID/*">
                        <xsl:variable name="TT" select="@TargetTag" />

<!-- Incluir na Saida ID mesmo que a Tag não exista -->
                        <xsl:variable name="TXNode" select="$Mensagem//*[name()=$TT]" />

                        <xsl:element name="{name()}">
                           <xsl:value-of select="LocalFunctions:flValorToSTR(string($TXNode), string(@Tipo), number(@Tamanho), number(@Decimais), string(@Obrigatorio) )" />
                        </xsl:element>
                     </xsl:for-each>
                  </xsl:element>
               </xsl:if>

<!-- Monta a Parte String-->
               <xsl:if test="$FormatoSTR">
                  <xsl:element name="SaidaSTR">
                     <xsl:for-each select="$FormatoSTR/*">
                        <xsl:variable name="TT" select="@TargetTag" />

                        <xsl:variable name="TXNode" select="$Mensagem//*[name()=$TT]" />

                        <xsl:element name="{name()}">
                           <xsl:value-of select="LocalFunctions:flValorToSTR(string($TXNode), string(@Tipo), number(@Tamanho), number(@Decimais), string(@Obrigatorio) )" />
                        </xsl:element>
                     </xsl:for-each>
                  </xsl:element>
               </xsl:if>

<!-- Monta a Parte XML-->
               <xsl:if test="$FormatoXML">
                  <xsl:element name="{$OutPutXMLName}">
                     <xsl:for-each select="$FormatoXML/*">
                        <xsl:variable name="TT" select="@TargetTag" />

                        <xsl:if test="$Mensagem//*[name()=$TT]">
                           <xsl:element name="{name()}">
                              <xsl:variable name="TXNode" select="$Mensagem//*[name()=$TT]" />

                              <xsl:variable name="TpNode" select="./@Tipo" />

                              <xsl:variable name="DecNode" select="./@Decimais" />

                              <xsl:variable name="ObriNode" select="./@Obrigatorio" />

                              <xsl:value-of select="LocalFunctions:flValorXMLparaXML(string($TXNode), string($TpNode),number($DecNode), string($ObriNode) )" />
                           </xsl:element>
                        </xsl:if>
                     </xsl:for-each>
                  </xsl:element>
               </xsl:if>

<!-- Monta a Parte CSV-->
               <xsl:if test="$FormatoCSV">
                  <xsl:element name="SaidaCSV">
                     <xsl:for-each select="$FormatoCSV/*">
                        <xsl:variable name="TT" select="@TargetTag" />

                        <xsl:if test="$Mensagem//*[name()=$TT]">
                           <xsl:element name="{name()}">
                              <xsl:value-of select="$Mensagem//*[name()=$TT]" />
                           </xsl:element>
                        </xsl:if>

                        <xsl:if test="position() != last()">
                           <xsl:element name="Separador">
                              <xsl:value-of select="../@Delimitador" />
                           </xsl:element>
                        </xsl:if>
                     </xsl:for-each>
                  </xsl:element>
               </xsl:if>
            </xsl:when>

<!--
***********************************************************************************************
  Entrada do tipo CSV - Separado por Delimitadores
***********************************************************************************************
-->
            <xsl:when test="//Documento/Mensagem/@Tipo='CSV'">
               <xsl:variable name="Delimitador" select="$Mensagem/@Delimitador" />

<!-- Monta a Parte ID -->
               <xsl:if test="$FormatoID">
                  <xsl:element name="SaidaID">
                     <xsl:for-each select="$FormatoID/*">
                        <xsl:variable name="CsvValue" select="LocalFunctions:flValueFromCSV(string($Mensagem), string($Delimitador), number(@Indice), string(@Tipo), boolean(@Obrigatorio))" />

                        <xsl:element name="{name()}">
                           <xsl:value-of select="LocalFunctions:flValorToSTR(string($CsvValue), string(@Tipo), number(@Tamanho), number(@Decimais), boolean(@Obrigatorio))" />
                        </xsl:element>
                     </xsl:for-each>
                  </xsl:element>
               </xsl:if>

<!-- Monta a Parte String -->
               <xsl:if test="$FormatoSTR">
                  <xsl:element name="SaidaSTR">
                     <xsl:for-each select="$FormatoSTR/*">
                        <xsl:variable name="CsvValue" select="LocalFunctions:flValueFromCSV(string($Mensagem), string($Delimitador), number(@Indice), string(@Tipo), boolean(@Obrigatorio))" />

                        <xsl:element name="{name()}">
                           <xsl:value-of select="LocalFunctions:flValorToSTR(string($CsvValue), string(@Tipo), number(@Tamanho), number(@Decimais), boolean(@Obrigatorio))" />
                        </xsl:element>
                     </xsl:for-each>
                  </xsl:element>
               </xsl:if>

<!-- Monta a Parte XML -->
               <xsl:if test="$FormatoXML">
                  <xsl:element name="{$OutPutXMLName}">
                     <xsl:for-each select="$FormatoXML/*">
                        <xsl:element name="{name()}">
                           <xsl:value-of select="LocalFunctions:flValueFromCSV(string($Mensagem), string($Delimitador), number(@Indice), string(@Tipo), boolean(@Obrigatorio))" />
                        </xsl:element>
                     </xsl:for-each>
                  </xsl:element>
               </xsl:if>

<!-- Monta a Parte CSV-->
               <xsl:if test="$FormatoCSV">
                  <xsl:element name="SaidaCSV">
                     <xsl:for-each select="$FormatoCSV/*">
                        <xsl:element name="{name()}">
                           <xsl:value-of select="LocalFunctions:flValueFromCSV(string($Mensagem), string($Delimitador), number(@Indice), string(@Tipo), boolean(@Obrigatorio))" />
                        </xsl:element>

                        <xsl:if test="position() != last()">
                           <xsl:element name="Separador">
                              <xsl:value-of select="../@Delimitador" />
                           </xsl:element>
                        </xsl:if>
                     </xsl:for-each>
                  </xsl:element>
               </xsl:if>
            </xsl:when>

            <xsl:otherwise>
               <Erro>Formato Não Reconhecido</Erro>

               <Formato>
                  <xsl:value-of select="//Documento/Mensagem/@Tipo" />
               </Formato>
            </xsl:otherwise>
         </xsl:choose>
      </Saida>
   </xsl:template>
</xsl:stylesheet>

