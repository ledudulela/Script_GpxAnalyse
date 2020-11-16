' Script "GPX Analyse"  
' --------------------  
' auteur: ledudulela  
' version 20201116.1440 
' màj: 16/11/2020   
' objectif: générère des statistiques à partir d'un fichier .gpx  
' HowTo: Copier le fichier gpx dans le répertoire du script. Lancer le script. Consulter les 3 fichiers générés.  
' Fichier.txt : statistiques globales (durée, vitesses, altitudes, distance)  
' Fichier.log : données gpx converties avec des unités de mesure usuelles (csv/tab)  
' Fichier.csv : données des points où dist_2pts > 0 + données calculées pour traitements avec un logiciel de type tableur (Excel, LibreOffice, etc)  
' le fichier .csv contient deux champs calculés: "Height" et "VSi" exploitable avec les graphique "Excel 3D-Map"  
' - height: hauteur du point par rapport à l'altitude la plus basse du parcours  
' - VSi (expérimental):  Vertical Speed Indicator dont les valeurs sont: -1=aucun, 0=faible, 1=moyen, 2=important  
' association conseillée de couleurs au VSi: -1:vert, 0:jaune, 1:rouge, 2:violet  
'  
' Altitudes corrigées : DEM (Digital Elevation Model) = MNA (modèle numérique d'altitude), de terrain ou de surface  
' --------------------  TED (Terrain Elevation Data)   
' En raison de l'imprécision des données d'altitudes des appareils enregistrant les données GPS en GPX, pour avoir des altitudes plus précises,  
' il est possible d'associer un fichier de données d'altitudes (DEM) en provenance du site www.gpsvisualizer.com/convert  
' Pour cela, dans le formulaire de convertion de gpsvisualizer,   
' - cocher Output Format = GPX  
' - sélectionner "Add DEM elevation data" = ODP1 - Western Europe     (ODP = https:data.opendataportal.at)  
' - uploader (téléverser) le fichier source GPX  
' Renommer le fichier obtenu comme le fichier source mais en remplacant l'extension .gpx par .ted (Terrain Elevation Data)  
' Puis placer le fichier .ted dans le même dossier que le fichier source  
' Le script vous indiquera la présence du fichier .ted et vous proposera de l'utiliser pour effectuer ses calculs.  
' Un nouveau fichier .gpx incluant les nouvelles données d'altitude sera alors créé si vous souhaitez utiliser les données "DEM".  
' PS: Le fichier obtenu par gpsvisualizer/convert perd les informations de vitesse du fichier gpx source.  
' Le script peut les fusionner pour en faire un nouveau fichier contenant les données GPX source + données d'altitudes provenant du TED  
  
' exemple de statistiques générées:	 
' ---------------------------------  
' Gpx-Filename=20201030_parcours_29km.gpx => nom du fichier gpx   
' Gpx-Creator=OsmAnd 3.8.3 => logiciel ayant enregistré la trace dans le fichier gpx  
' Gpx-Name=2020-10-25_09-03_Sun => nom interne du fichier gpx  
' TrackPoints=719 (moving) / 829 => nbr de points enregistrés dans le fichier (avec le nbr de points où la vitesse > 0)  
' AvgRecTimer=10.6 s => fréquence à laquelle le logiciel a enregistré chaque point dans le fichier  
' Date=2020-10-25 => date de création du fichier  
' StartTime=08:03:01 => heure de début de l'enregistrement  
' EndTime=10:29:26 => heure de fin de l'enregistrement  
' TotalTime=2 h 26 min => durée du parcours  
' MovingTime=2 h 7 min => durée du parcours en mouvement (vitesse>0)  
' TotalDistance=28.626 km => distance totale du parcours  
' MaxSpeed=31.6 km/h => vitesse maximum  
' AvgSpeed=13.5 km/h (moving) => moyenne de vitesse calculée sur les points dont la vitesse > 0 (pour exclure les pauses)  
' MinElevation=54.6 m => altitude la plus basse du parcours  
' MaxElevation=77.1 m => altitude la plus élevée du parcours  
' DiffElevation=22.5 m => différence entre altitude min et altitude max  
' TotalAscent=457 m => cumul des montées  
' TotalDescent=462 m => cumul descentes  
' ```  
' exemple de contenu de fichier GPX généré par Osmand: (il peut exister des variantes selon l'origine du GPX, en particulier pour le champ "speed" )  
' ----------------------------------------------------  
' <trkpt lat="44.0" lon="-0.5">  
'  <ele>18.9</ele>  
'  <time>2020-09-29T15:47:54Z</time>  
'  <hdop>1.6</hdop>  
'  <extensions>  
'    <speed>0.324</speed>  
'  </extensions>  
' </trkpt>  
' ```
option explicit
const VSI_OFFSET=0 ' ajustement des plages de valeurs pour le calcul du VSI, en centimetres/sec ! exemple: VSI_OFFSET=9 pour vttiste confirmé
const TITLE="GPX Analyse"
dim CSV
dim CRLF
CSV=chr(9)
CRLF=chr(13) & chr(10)

Function CurrentDir() ' répertoire en cours
	dim monRepScript
	monRepScript=left(wscript.scriptfullname,len(wscript.scriptfullname)-len(wscript.scriptname))
	CurrentDir=monRepScript
End Function
		
sub main () ' choix du fichier gpx et execution de l analyse
	const GPX=".gpx"
	dim strFileName
	dim arrFiles
	dim strChoix
	dim i
	Dim fso, f, f1, fc, strFileList

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFolder(CurrentDir())
	Set fc = f.Files
	
	' liste des fichiers du répertoire en cours
	strFileName=""
	strFileList=""
	For Each f1 in fc
		if instr(lcase(f1.name),GPX)>0 then
			if right(lcase(f1.name),4)=GPX then
				i=i+1
				strFileList = strFileList & i & ". " & f1.name 
				strFileList = strFileList & CRLF
			end if
		end if
	Next
	
	' choix du fichier parmi la liste des fichiers du répertoire en cours
	i=0
	if strFileList<>"" then
		arrFiles=split(strFileList,CRLF)
		strChoix=inputbox(strFileList & CRLF & "Entrez le numero du fichier:",TITLE,1) 'replace(wscript.scriptname,".vbs",".gpx")
		if strChoix<>"" then
			if isnumeric(strChoix) then
				i=int(strChoix)-1
				if i>=lbound(arrFiles) and i<ubound(arrFiles) then
					strChoix=split(arrFiles(i),". ")
					strFileName=strChoix(1)
				end if
			end if
		end if
	else
		msgbox "Ce dossier ne contient aucun fichier .gpx .",,TITLE
	end if
	
	' lance l'analyse le fichier GPX
	if strFileName<>"" then
		gpxAnalyse CurrentDir() & strFileName,1,0,0
	end if

end sub

sub gpxAnalyse(paramFileFullPath,paramPass, paramMinElevM, paramMovingPts)
	' La procedure est executee deux fois. 
	' La premiere, permet de recuperer l altitude mini sur le parcours et éventuellement de mapper les altitudes DEM.
	' La seconde, permet de calculer la hauteur des points par rapport à cette altitude mini,
	' calcul: Hauteur = ElevationM - minElevM
	' Exemple: si altitude mini est 40m et qu'un point est a 55m alors la hauteur du point est (55-40)=15m
	' Cela permet, sur les graphiques de bien voir les differences d'altitude entre chaque point.
	' PS: les resultats sont ecrits dans les fichiers (.txt .log .csv) lors de la seconde passe.
	const boolDEBUG=false
	
	const vbYesNo=4 
	const vbYes=6
	const vbNo=7
	
	const EXPE_VALUES=false
	const XML_DOM_DOC="MSXML2.DOMDocument"
	const umS=" s" ' seconde
	const umM=" m" ' metre
	const umKM=" km"
	const umKMH=" km/h"
	const umNM=" NM" 'nautical miles
	
	dim stdOut
	dim objFSO
	dim oLogFile
	dim oCSVFile
	dim oFile
	dim objADODB
	
	dim strFileContent
	dim strFileFullPath
	dim boolWriteFile
	
	dim xmlDOMSource
	dim xmlDomTED
	dim xmlRootNode
	dim xmlMetadataNode
	dim xmlRows
	dim xmlRowsTED
	dim xmlRow
	dim xmlRowTED
	dim xmlPreviousRow
	dim strRow
	dim l
	
	' *** objets xml ***
	dim gpxMetadataName
	dim gpxAttrLat
	dim gpxAttrLon
	dim gpxEle
	dim gpxTime
	dim gpxHdop
	dim gpxExtensions
	dim gpxDesc
	dim gpxSpeed
	dim gpxEleTED
	
	' *** valeur des objets XML *** 
	dim ptLat
	dim ptLon
	dim ptElevationM
	dim ptDateTime
	dim ptDateOnly
	dim ptHdop	
	dim ptSpeedMS
	
	' *** valeurs à afficher ***
	dim varGpxFilename
	dim varCreator
	dim varGpxName
	dim varMinDateTime
	dim varMaxDateTime
	dim varMinDate
	'dim varMaxDate
	dim varRecTimerS
	dim varTrackpoints
	dim varDate
	dim varStartTime
	dim varEndTime
	dim varTotalTime
	dim varTotalDistanceM
	dim varMinElevation
	dim varMaxElevation
	dim varDiffElevation
	dim varTotalDescent
	dim varTotalAscent
	dim varMovingTime
	dim varMovingPercent
	dim varMovingDistM
	dim varMaxSpeedMS
	
	' *** pour les calculs ***
	dim ptPreviousLat
	dim ptPreviousLon
	dim ptPreviousElevationM
	
	dim dblSpeedMS
	
	dim segDistanceM
	dim segHeightM
	dim segHillSign
	dim segAscend
	dim segDescend
	
	dim intDiffTS
	dim nbrMovingPt
	' dim isMovingPt
	dim previousMovingEle
	dim previousMovingDateTime
	dim startMovingDateTime
	dim intDurationS
	dim dblVerticalSpeed
	dim intVSi
	dim arrDesc
	dim intDiffTime
	
	dim boolAscent
	dim sumAscent
	dim maxAscent
	dim sumDistAscent
	dim maxDistAscent
	
	dim boolDescent
	dim sumDescent
	dim maxDescent
	dim dblHeightM
	
	dim boolWithTED
	dim boolErrTED
	
	dim boolGpxTS
	dim newGpxTS
	dim strNewGPXTime
	'dim x
	'dim y
	'dim z
	'dim radius
	
	' **************************************
	' l'objet FileSystemObject va être utilisé à plusieurs reprises dans le code
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	if boolDEBUG then
		Set stdOut=objFSO.getStandardStream(1) ' compatible cscript.exe windows 10, pour afficher la trace dans la console 
	end if
	
	' récupère le contenu du fichier.gpx 
	strFileContent=""
	if objFSO.FileExists(paramFileFullPath) then 'strFileFullPath=CurrentDir() & strFileName 
		Set objADODB = CreateObject("ADODB.Stream") ' cet objet permet de convertir un fichier encodé en UTF8 en une chaine ANSI
		objADODB.Charset="utf-8"
		objADODB.open
		objADODB.loadFromFile(paramFileFullPath)
		strFileContent=objADODB.ReadText()
		objADODB.close
		set objADODB=nothing
	end if

	if strFileContent="" then 
		msgbox "Erreur de fichier source.",,TITLE
	else	
		varGpxFilename=objFSO.getFileName(paramFileFullPath)
		boolWriteFile=(paramPass=2) ' booleen pour indiquer d ecrire dans les fichiers à la seconde passe de la fonction
		
		' charge le XML du fichier GPX
		set xmlDOMSource = CreateObject( XML_DOM_DOC )
		xmlDOMSource.loadXML(strFileContent) 
		
		' verifie la presence du fichier GPX avec DEM et le charge
		boolWithTED=false
		if paramPass=1 then
			strFileFullPath=replace(paramFileFullPath,".gpx",".ted")
			strFileContent=""
			if objFSO.FileExists(strFileFullPath) then 
				Set objADODB = CreateObject("ADODB.Stream") ' cet objet permet de convertir un fichier encodé en UTF8 en une chaine ANSI
				objADODB.Charset="utf-8"
				objADODB.open
				objADODB.loadFromFile(strFileFullPath)
				strFileContent=objADODB.ReadText()
				objADODB.close
				set objADODB=nothing
				
				set xmlDomTED = CreateObject( XML_DOM_DOC )
				xmlDomTED.loadXML(strFileContent) ' charge le XML	
				set xmlRootNode = xmlDomTED.selectSingleNode( "//gpx" )	' sélectionne le noeud racine
				set xmlRowsTED=xmlDomTED.getElementsByTagName("trkpt") ' collection des éléments trkpt
				boolWithTED=true
			end if
		end if		
		
		' recuperation de metadonnees
		varCreator=""
		set xmlRootNode = xmlDOMSource.selectSingleNode( "//gpx" )	' sélectionne le noeud racine
		varCreator=xmlRootNode.attributes.getNamedItem("creator").value
		
		varTrackpoints=0
		set xmlRows=xmlDOMSource.getElementsByTagName("trkpt") ' collection des éléments trkpt
		varTrackpoints=xmlRows.length
		
		varGpxName=""
		set xmlMetadataNode = xmlRootNode.selectSingleNode("metadata")	' sélectionne le noeud metadata
		if not xmlMetadataNode is nothing then
			set gpxMetadataName=xmlMetadataNode.selectSingleNode("name")
			if not gpxMetadataName is nothing then	
				varGpxName=gpxMetadataName.text
			end if
		end if
		
		' initialisation des variables
		l=0
		ptSpeedMS=""
		ptPreviousLat=""
		ptPreviousLon=""
		ptPreviousElevationM=""
		
		varStartTime=""
		varEndTime=""
		varTotalTime=0
		varTotalDistanceM=0
		varMinElevation=9999
		varMaxElevation=0	
		varDiffElevation=0
		varMaxSpeedMS=0
		varTotalAscent=0
		varTotalDescent=0
		varMovingDistM=0
		varMovingTime=0
		varMovingPercent=0
		nbrMovingPt=0
		startMovingDateTime=""
		previousMovingEle=""
		previousMovingDateTime=""
		intDurationS=0
		dblVerticalSpeed=0
		intVSi=0
		
		' expérimental
		boolAscent=false
		sumAscent=0
		maxAscent=0
		boolDescent=false
		sumDescent=0
		maxDescent=0
		
		' Si fichier DEM présent alors remplace les donnees d'elevation du fichier source par les donnees d'elevation du fichier DEM
		' et enregistre les nouvelles données gpx dans un nouveau fichier  "nomfichiersource_tex.gpx"
		' Si le fichier GPX source ne contient pas de time ou de vitesse, en ajoute un dans le nouveau fichier.
		boolErrTED=false
		boolGpxTS=false
		if boolWithTED then
			newGpxTS=toSystemTS("1901-12-31T00:00:00Z")
			if xmlRows.length=xmlRowsTED.length then ' on verifie que le fichier TED contient le meme nbr de points que le fichier GPX source
				strFileFullPath=replace(paramFileFullPath,".gpx","_TED.gpx") ' nom du nouveau fichier cible incluant les "Terrain Elevation Data"
				if msgbox ("Voulez-vous utiliser les altitudes du fichier TED" & _
					" et creer le fichier: " & objFSO.getFileName(strFileFullPath) & " ?",vbYesNo,TITLE)=vbYes then
					for each xmlRow in xmlRows	
						set xmlRowTED=xmlRowsTED(l) ' la meme ligne que le source
						
						' *** verifie que les points Source et DEM ont bien les memes coordonnees
						' attention contrairement à OsmAnd, STRAVA ajoute des 0 non significatifs aux coordonnees.
						if RoundDown(cDouble(xmlRow.attributes.getNamedItem("lat").value),4) = RoundDown(cDouble(xmlRowTED.attributes.getNamedItem("lat").value),4) and _
							RoundDown(cDouble(xmlRow.attributes.getNamedItem("lon").value),4)=RoundDown(cDouble(xmlRowTED.attributes.getNamedItem("lon").value),4) then
							set gpxEle=xmlRow.selectSingleNode("ele")
							set gpxEleTED=xmlRowTED.selectSingleNode("ele")
							
							' *** ajoute une elevation si n existe pas
							if gpxEle is nothing then 
								set gpxEle = xmlDOMSource.createElement("ele")
								xmlRow.appendChild(gpxEle)
							end if
							gpxEle.text=gpxEleTED.text ' nouvelle valeur de Elevation from DEM
							'stdOut.WriteLine l & gpxEle.text & " " & gpxEleTED.text
							
							' *** ajoute un Time bidon si n existe pas dans le fichier source
							varRecTimerS=10 ' secondes, valeur arbitraire mais volontairement elevee pour avoir une vitesse faible
							set gpxTime=xmlRow.selectSingleNode("time")
							if  gpxTime is nothing then
								
								if l=0 then
									if msgbox ("Le fichier ne contient pas de donnees temporelles, voulez-vous en ajouter des fictives ?",vbYesNo,TITLE)=vbYes then
										boolGpxTS=true
									end if
								end if
								
								if boolGpxTS then
									newGpxTS=dateadd("s",varRecTimerS,newGpxTS)
									strNewGPXTime=datepart("yyyy",newGpxTS) & "-" & T2digits(datepart("m",newGpxTS)) & "-" & T2digits(datepart("d",newGpxTS)) & "T" & _
									T2digits(datepart("h",newGpxTS)) & ":" & T2digits(datepart("n",newGpxTS)) & ":" & T2digits(datepart("s",newGpxTS)) & "Z"
								
									set gpxTime = xmlDOMSource.createElement("time")
									gpxTime.text=strNewGPXTime
									xmlRow.appendChild(gpxTime)
									
									' ci-dessous ça marchait mais je pense inutile
									'
'									' *** ajoute une vitesse fictive si n existe pas dans le fichier source 
'									set gpxExtensions=xmlRow.selectSingleNode("extensions")
'									if  gpxExtensions is nothing then
'										set gpxExtensions = xmlDOMSource.createElement("extensions")
'										set gpxSpeed = xmlDOMSource.createElement("speed")
'										segDistanceM=0
'										segHeightM=0
'										if l>0 then
'											' calcul de vitesse par rapport au point précédent
'											set xmlPreviousRow=xmlRows(l-1)
'											segHeightM=cDouble(xmlRow.selectSingleNode("ele").text) - cDouble(xmlPreviousRow.selectSingleNode("ele").text)
'											segDistanceM=distance2ptsM( _
'												xmlRow.attributes.getNamedItem("lat").value, _
'												xmlRow.attributes.getNamedItem("lon").value, _
'												xmlPreviousRow.attributes.getNamedItem("lat").value, _
'												xmlPreviousRow.attributes.getNamedItem("lon").value, _
'												segHeightM)
'											ptSpeedMS=Dot(RoundDown(segDistanceM/varRecTimerS,1)) ' distance / temps
'											'stdOut.writeLine l & " " & ptSpeedMS & "m/s"
'										else
'											ptSpeedMS="0.0"
'										end if
'										gpxSpeed.text=ptSpeedMS
'										gpxExtensions.appendChild(gpxSpeed)
'										xmlRow.appendChild(gpxExtensions)
'									end if
								end if
							end if
						else
							boolErrTED=true
							exit for
						end if
						l=l+1
					next
					
					if boolErrTED then ' normalement ce type d erreur ne devrait pas arriver
						msgbox "TED: Erreur de correspondance de coordonnees au point " & l,,TITLE
					else
						paramFileFullPath=strFileFullPath ' on memorise le chemin pour la seconde passe
						xmlDOMSource.save paramFileFullPath ' sauvegarde le nouveau fichier gpx avec altitudes corrigees
						if ptSpeedMS<>"" then
							msgbox "Des donnees fictives de temps (" & varRecTimerS & " sec) ont ete ajoutees au fichier TED nouvellement cree.",,TITLE
						end if
					end if
				
				end if
			
			else ' ca ne devrait pas arriver
				msgbox "Probleme, le fichier TED ne contient pas le meme nbr de points que le fichier GPX source.",,TITLE
			end if
		end if
		
		' ----------------------------------------------------------------------------------------
		'  c est parti pour l analyse...
		' ----------------------------------------------------------------------------------------
		l=0
		strFileFullPath=CurrentDir() & replace(wscript.scriptname,".vbs",".log")
		Set oLogFile=objFSO.OpenTextFile(strFileFullPath,2,true,0) ' ecriture,create, ascii
		
		strFileFullPath=CurrentDir() & replace(wscript.scriptname,".vbs",".csv")
		Set oCSVFile=objFSO.OpenTextFile(strFileFullPath,2,true,0) ' ecriture,create, ascii
		if boolWriteFile then ' ecrit la premiere ligne du fichier CSV (entêtes de colonnes)
			strRow=	"Pt" & CSV & "Time" & CSV & "Duration" & CSV & "DistKM" & CSV & _
					"SpeedKMH" & CSV & "ElevM" & CSV & "DiffM" & CSV & _
					"Lon" & CSV & "Lat" & CSV & "Height" & CSV & "VSI"
			oCSVFile.writeLine strRow
		end if
		
		' boucle sur chaque elt xml
		for each xmlRow in xmlRows	
			l=l+1
			
			set gpxAttrLat=xmlRow.attributes.getNamedItem("lat")
			set gpxAttrLon=xmlRow.attributes.getNamedItem("lon")
			set gpxEle=xmlRow.selectSingleNode("ele")
			set gpxTime=xmlRow.selectSingleNode("time")
			set gpxHdop=xmlRow.selectSingleNode("hdop")

			if gpxAttrLat is nothing or _
				gpxAttrLon is nothing or _
				gpxEle is nothing or _
				gpxTime is nothing then 
				' on ne fait rien si ces elts ne sont pas trouvés.
			else
				' mémorise les valeurs des elts
				ptLat=gpxAttrLat.value
				ptLon=gpxAttrLon.value
				ptElevationM=gpxEle.text
				ptDateTime=gpxTime.text
				ptDateOnly=extractDate(gpxTime.text)
				'ptHdop=gpxHdop.text
				
				if l=1 then ' pour le permier point, initialisation de variables
					varMinDateTime=ptDateTime
					varMaxDateTime=ptDateTime
					varMinDate=ptDateOnly
					'varMaxDate=ptDateOnly
					previousMovingDateTime=ptDateTime
				end if
				
				if ptDateOnly<varMinDate then  ' 2020-09-29  ( T15:47:54Z )
					varMinDate=ptDateOnly ' si la trace est à cheval sur 2 jours, on ne garde que la date min
				end if
				
				'if ptDateOnly>varMaxDate then  ' 2020-09-29  ( T16:55:23Z )
				'	varMaxDate=ptDateOnly
				'end if
				
				' dans tous les cas, le max DateTime est le nouveau DateTime
				varMaxDateTime=ptDateTime				

				
				' vitesse
				ptSpeedMS=""
				' isMovingPt=false
				
				' vitesse dans un GPX OsmAnd
				set gpxExtensions=xmlRow.selectSingleNode("extensions")
				if not gpxExtensions is nothing then
					set gpxSpeed=gpxExtensions.selectSingleNode("speed")
					if not gpxSpeed is nothing then	
						ptSpeedMS=trim(gpxSpeed.text)
						dblSpeedMS=cDouble(ptSpeedMS)
					end if
				end if
				
				' vitesse dans un GPX GPS-Visualyser/Export
				set gpxSpeed=xmlRow.selectSingleNode("speed")
				if not gpxSpeed is nothing then
					ptSpeedMS=trim(gpxSpeed.text)
					dblSpeedMS=cDouble(ptSpeedMS)				
				end if
				
				' vitesse dans un GPX Sport-Tracker (version 2011 Symbian - Terence)
				' <desc>Speed 6.3 km/h Distance 0.00 km</desc>
				set gpxDesc=xmlRow.selectSingleNode("desc")
				if not gpxDesc is nothing then
					if instr(gpxDesc.text,"Speed ") then
						arrDesc=split(gpxDesc.text," ")
						dblSpeedMS=cDouble(arrDesc(1))
						if arrDesc(2)="km/h" then dblSpeedMS=dblSpeedMS*1000/3600
						ptSpeedMS=dot(dblSpeedMS)
					end if
				end if
				
				' mémorise la Vitesse Max
				if ptSpeedMS<>"" then
					if dblSpeedMS>0 then ' moving point
						if dblSpeedMS>varMaxSpeedMS then
							varMaxSpeedMS=dblSpeedMS
						end if
					end if						
				end if
				
				' calcul elevation / pente
				segHeightM=0
				segHillSign="" 
				if ptElevationM<>"" then
					if cDouble(ptElevationM)>varMaxElevation then
						varMaxElevation=cDouble(ptElevationM) ' mémorise la nouvelle Altitude max
					end if
					if cDouble(ptElevationM)<varMinElevation then
						varMinElevation=cDouble(ptElevationM) ' mémorise la nouvelle Altitude min
					end if
					if ptPreviousElevationM<>"" then
						' calcul de la pente entre le point en cours et le point précédent (différence d'altitude)
						segHeightM=cDouble(ptElevationM)-cDouble(ptPreviousElevationM)
						if segHeightM>0 then
							varTotalAscent=varTotalAscent+segHeightM
							segHillSign="+" ' montée
						else
							varTotalDescent=varTotalDescent+segHeightM
							segHillSign="-" ' descente
						end if

					end if
				end if
				ptPreviousElevationM=ptElevationM ' mémorise l'altitude du point en cours pour la prochaine entrée dans la boucle
				
				' calcul de distance entre le point et le point précédent
				segDistanceM=0
				if ptLat<>"" and ptLon<>"" and ptPreviousLat<>"" and ptPreviousLon<>"" then
					segDistanceM=distance2ptsM(ptLat,ptLon,ptPreviousLat,ptPreviousLon,segHeightM)
				end if			
				ptPreviousLat=ptLat
				ptPreviousLon=ptLon
				
				' calcul TotalDistanceM
				varTotalDistanceM = varTotalDistanceM + segDistanceM				
				
				' construit la ligne à écrire dans le fichier .log
				if boolWriteFile then 
					strRow=LblValUnitCsv("ts",extractTime(ptDateTime),"") & _ 
						LblValUnitCsv("lat",ptLat,"") & _
						LblValUnitCsv("lon",ptLon,"") & _
						LblValUnitCsv("speed",ptSpeedMS,"m/s") & _
						LblValUnitCsv("dist",RoundUp(segDistanceM,2),"m") & _
						LblValUnitCsv("elev",ptElevationM,"m") & _
						LblValUnitCsv("diff",segHillSign & RoundUp(abs(segHeightM),2),"m") 
					oLogFile.writeLine strRow 
				end if
				'stdOut.WriteLine strRow
				

				
				' pour les stats sur les points en mouvement uniquement (vers le fichier .csv)
				if segDistanceM > 0 then ' isMovingPt
					nbrMovingPt=nbrMovingPt+1 ' comptabilise le nbr de pts en mouvement
					
					' si le GPX ne contient pas de Speed, alors speed=0 
					if ptSpeedMS="" then ptSpeedMS="0.0" 
					
					if startMovingDateTime="" then ' initialisation
						startMovingDateTime=ptDateTime
						previousMovingEle=ptElevationM	
					else
						' durée depuis le start_Moving_DateTime
						intDurationS=datediff("s",toSystemTS(startMovingDateTime),toSystemTS(ptDateTime)) 				
					end if
					
					' cumul de longueur des segments
					varMovingDistM=varMovingDistM + segDistanceM
					
					' hauteur du point par rapport a l'altitude minimum du parcours
					dblHeightM=cDouble(ptElevationM)-paramMinElevM
					
					' pente du segment (hauteur)
					segHeightM=RoundDown(cDouble(ptElevationM)-cDouble(previousMovingEle),1)
					
					' Calcul de "difficulté de montée"
					' Ces valeurs arbitraires semblent bien adaptées au VTT
					' On considère que plus on met de temps pour s'élever d'un mètre (par exemple), plus la difficulté est importante.
					' hauteur du segment / durée du segment
					intDiffTime=datediff("s",toSystemTS(previousMovingDateTime),toSystemTS(ptDateTime))
					if intDiffTime=0 then intDiffTime=1 ' stdOut.WriteLine l & "_" & previousMovingDateTime & "_" & ptDateTime & "_"
					dblVerticalSpeed=segHeightM*100 / intDiffTime ' mètre => cm
					
					intVSi=0 ' VSi = indicateur de vitesse verticale
					'les valeurs étant arbitraire, on peut les ajuster avec la constante VSI_OFFSET (VerticalSpeed en cm)
					if dblVerticalSpeed<0 then intVSi=-1
					if dblVerticalSpeed>=0 and dblVerticalSpeed<(10 + VSI_OFFSET) then intVSi=0
					if dblVerticalSpeed>=(10 + VSI_OFFSET) and dblVerticalSpeed<(20 + VSI_OFFSET) then intVSi=1
					if dblVerticalSpeed>=(20 + VSI_OFFSET) then intVSi=2
					
					' ecriture de la ligne dans le fichier CSV 
					if boolWriteFile then 
						strRow=l & CSV & _
							extractTime(ptDateTime) & CSV & _
							secToHMS(intDurationS) & CSV & _
							RoundUp(varMovingDistM/1000,3) & CSV & _
							RoundUp(MStoKMH(cDouble(ptSpeedMS)),1) & CSV & _
							cDouble(ptElevationM) & CSV & _
							segHeightM & CSV & _
							cDouble(ptLon) & CSV & cDouble(ptLat) & CSV & _
							RoundUp(dblHeightM,3) & CSV & _
							intVSi
						oCSVFile.WriteLine strRow
					end if
					
					' pour calcul diff altitude
					previousMovingEle=ptElevationM
					previousMovingDateTime=ptDateTime
				end if ' fin de traitement du "moving point"
				
			end if ' track is nothing
		next
		' apres la boucle, ferme les fichiers
		oLogFile.Close	
		oCSVFile.Close	
		
		' calculs et ecriture des statistiques globales dans le fichier texte
		varDate=varMinDate
		varStartTime=extractTime(varMinDateTime)
		varEndTime=extractTime(varMaxDateTime)
		varTotalTime=datediff("s",toSystemTS(varMinDateTime),toSystemTS(varMaxDateTime)) 
		varRecTimerS=RoundDown(varTotalTime/varTrackpoints,1)
		varMovingTime=nbrMovingPt*varRecTimerS ' nbr_pt_en_mouvement x rec_timer
		if varMovingTime=0 then varMovingTime=varTotalTime
		if varMovingTime=0 then varMovingTime=1 ' pour eviter la division par 0
		if varTrackpoints>0 then varMovingPercent=RoundUp(nbrMovingPt / varTrackpoints * 100,1) 
		varMinElevation=RoundDown(varMinElevation,1)
		varMaxElevation=RoundDown(varMaxElevation,1)
		varDiffElevation=RoundDown(varMaxElevation-varMinElevation,1)
		if boolWriteFile then 	
			strFileContent=LblVal("Gpx-Filename", varGpxFilename) ' & nom du gpx
			strFileContent=strFileContent & CRLF & LblVal("Gpx-Creator",varCreator)
			strFileContent=strFileContent & CRLF & LblVal("Gpx-Name",varGpxName)
			strFileContent=strFileContent & CRLF & LblVal("TrackPoints", nbrMovingPt & " (moving) / " & varTrackpoints) & " (" & varMovingPercent & "%)"
			strFileContent=strFileContent & CRLF & LblVal("AvgRecTimer", varRecTimerS & umS)
			strFileContent=strFileContent & CRLF & LblVal("Date", varDate)
			strFileContent=strFileContent & CRLF & LblVal("StartTime", varStartTime)
			strFileContent=strFileContent & CRLF & LblVal("EndTime", varEndTime)
			strFileContent=strFileContent & CRLF & LblVal("TotalTime", secToHMin(varTotalTime))
			strFileContent=strFileContent & CRLF & LblVal("MovingTime", secToHMin(varMovingTime))
			strFileContent=strFileContent & CRLF & LblVal("TotalDistance", RoundUp(varTotalDistanceM/1000,3) & umKM)
			strFileContent=strFileContent & CRLF & LblVal("MaxSpeed",RoundUp(MStoKMH(varMaxSpeedMS),1) & umKMH)
			strFileContent=strFileContent & CRLF & LblVal("AvgSpeed", RoundUp(varTotalDistanceM/varMovingTime*3600/1000,1) & umKMH & " (moving)")
			strFileContent=strFileContent & CRLF & LblVal("MinElevation", varMinElevation & umM)
			strFileContent=strFileContent & CRLF & LblVal("MaxElevation", varMaxElevation & umM)
			strFileContent=strFileContent & CRLF & LblVal("DiffElevation", varDiffElevation & umM)
			strFileContent=strFileContent & CRLF & LblVal("TotalAscent", int(varTotalAscent) & umM)
			strFileContent=strFileContent & CRLF & LblVal("TotalDescent", int(varTotalDescent*-1) & umM)
			
			if EXPE_VALUES then
				'strFileContent=strFileContent & CRLF & LblVal("MinDate", varMinDate)
				'strFileContent=strFileContent & CRLF & LblVal("MaxDate", varMaxDate)
				strFileContent=strFileContent & CRLF & LblVal("MaxAscent", int(maxAscent) & umM & " sur " & int(maxDistAscent) & umM & " " & int(maxAscent/maxDistAscent*100) & "%")
				strFileContent=strFileContent & CRLF & LblVal("MaxDescent", int(maxDescent) & umM)
			end if
			
			' ecriture des statistiques dans le fichier texte
			strFileFullPath=CurrentDir() & replace(wscript.scriptname,".vbs",".txt")
			Set oFile=objFSO.OpenTextFile(strFileFullPath,2,true,0) ' ecriture,create, ascii
			oFile.Write strFileContent
			oFile.Close
		end if
		
		if paramPass=1 then
			if varTrackpoints=0 then
				msgbox "Le fichier ne contient pas de donnees analysables.",,TITLE
			else
				' Relance l'analyse avec cette fois-ci le MinElevation en parametre pour calculer 
				' les hauteurs des points par rapport au MinElevation.
				
				gpxAnalyse paramFileFullPath,2,varMinElevation,nbrMovingPt
			end if
		end if
		
		if boolWriteFile then ' normalement a la seconde passe		
			msgbox strFileContent & CRLF & CRLF & "Les donnees ont ete sauvegardees dans les fichiers .txt .log et .csv .",,TITLE
		end if
	
	end if
end sub

' --------------------------------------------------------------------------------------------------------------------------------------
' quelques fonctions utiles
' --------------------------------------------------------------------------------------------------------------------------------------
function dot(value) ' convertit la virgule en point 
	dot=replace(value,",",".")
end function

function LblVal(label, value) ' chaine concaténée libellé + valeur
	LblVal=label & "=" & dot(value)
end function

function LblValUnitCsv(label, value, unit) ' chaine concaténée libellé + valeur + unité de mesure + séparateur CSV
	LblValUnitCsv=label & "=" & dot(value) & unit & CSV
end function

function MStoKMH (varSpeedMS) ' conversion mètre/sec en km/h
	MStoKMH=varSpeedMS*3600/1000
end function

function cDouble(strNum) ' renvoie un double
	' si c'est un point en guise de séparateur de décimal, alors convertit en replace le point par une virgule et renvoie un double
	cDouble=cdbl(replace(strNum,".",","))
end function

function extractDate(strDateTime) ' 2020-12-31
	extractDate=left(strDateTime,10)  ' partie gauche de la chaine DateTime (2020-12-31T12:30:59Z)
end function

function extractTime(strDateTime) ' 12:30:59
	extractTime=mid(strDateTime,12,8) ' partie droite de la chaine DateTime (2020-12-31T12:30:59Z)
end function

function toSystemTS(strDateTime) ' renvoie une date système (vba)
	dim myDateTime
	myDateTime=trim(replace(left(strDateTime,19),"T"," "))
	if strDateTime<>"" then
		toSystemTS=cdate(myDateTime) ' convertir la chaine "DateTime" en une variable de type "Date système" (timestamp)
	else
		toSystemTS=date
	end if
end function

function secToHMin(intSec) ' convertit des secondes en Heure + Min
	dim h 
	dim m
	dim intMin
	dim strReturn
	intMin=int(intSec/60)
	h=int(intMin/60)
	m=intMin-(h*60)
	strReturn=""
	if h>0 then strReturn=h & " h " 
	strReturn=strReturn & m & " min"
	secToHMin=strReturn
end function

function secToHMS(intSec) ' convertit des secondes en Heure + Min + Sec, au format 12:45:59
	dim h 
	dim m
	dim s
	dim intMin
	dim strReturn
	intMin=int(intSec/60)
	h=int(intMin/60)
	m=intMin-(h*60)
	s=intSec -(h*3600) -(m*60)
	
	strReturn=""
	strReturn=T2digits(h) & ":" ' h
	strReturn=strReturn & T2digits(m) & ":" ' m
	strReturn=strReturn & T2digits(s) ' & "s"
	secToHMS=strReturn
end function

function T2digits(intValue) ' un chiffre au format 00. exempe: 2 -> 02
	T2digits=right("0" & intValue,2)
end function

Function RoundDown(value, digits) ' fonction d'arrondi
    RoundDown = Int((value + (1 / (10 ^ (digits + 1)))) * (10 ^ digits)) / (10 ^ digits)
End Function

Function RoundUp(value, digits) ' fonction d'arrondi 
    RoundUp = RoundDown(value + (5 / (10 ^ (digits + 1))), digits)
End Function

Function toRadians(value) ' convertit en radian
 dim PI
 PI = Atn(1) * 4
 toRadians = value * PI / 180
End Function

function distance2ptsM(strLatDepart,strLonDepart,strLatArrivee,strLonArrivee,height) ' distance entre 2 points, en tenant compte de la diférence d'altitude entre les 2 points
	dim distNM
	dim distM
	distNM=distance2ptsNM(strLatDepart,strLonDepart,strLatArrivee,strLonArrivee)
	distM=distNM*1852
	distM=sqr(distM*distM+height*height) ' pythagore, sur un triagle rectangle c²=a²+b² où distM = C
	distance2ptsM=distM
end function

function distance2ptsNM(strLatDepart,strLonDepart,strLatArrivee,strLonArrivee) ' distance entre deux positions géographiques satellite
	dim result
	dim tempo
	dim latDepart
	dim lonDepart
	dim latArrivee
	dim lonArrivee
	latDepart=cDouble(strLatDepart)
	lonDepart=cDouble(strLonDepart)
	latArrivee=cDouble(strLatArrivee)
	lonArrivee=cDouble(strLonArrivee)
	
	result=0
	if latDepart=latArrivee then
		result=60 * (abs(latDepart-latArrivee)) * cos(toRadians(latDepart))
	else
		tempo=atn((abs(lonDepart-lonArrivee)) / (abs(latDepart-latArrivee)) * cos(toRadians((abs(latDepart + latArrivee)) /2)))
		result=(abs(latDepart - latArrivee)) / cos(tempo) * 60
	end if
	
	distance2ptsNM=result ' resultat calcul en NM 
end function

sub LatLonToXYZ(dblLat, dblLon, dblRadius, dblX, dblY, dblZ) ' inutilisée
	dblY=dblRadius * sin(toRadians(dblLat))
	dblX=dblRadius * sin(toRadians(dblLon)) * cos(toRadians(dblLat))
	dblZ=-dblRadius * cos(toRadians(dblLon)) * cos(toRadians(dblLat))
end sub


' -----------------------------------------------
' excéute la fonction principale du script
' -----------------------------------------------
main