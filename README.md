# Script_GpxAnalyse
Doc màj: 22/11/2020   

objectif: générer des statistiques à partir d'un fichier .gpx  

![screenshot](https://github.com/ledudulela/Script_GpxAnalyse/blob/main/gpxanalyse01.jpg)

![screenshot](https://github.com/ledudulela/Script_GpxAnalyse/blob/main/gpxanalyse02.jpg)

![screenshot](https://github.com/ledudulela/Script_GpxAnalyse/blob/main/gpxanalyse03.jpg)

Le script est téléchargeable ici: http://ledudulela.free.fr/public/scripts/GPX/  

HowTo: Copier le fichier gpx dans le répertoire où se trouve le script. Lancer le script. Consulter les 3 fichiers générés.  

Fichier.txt : statistiques globales (durée, vitesses, altitudes, distance)  

Fichier.log : données gpx converties avec des unités de mesure usuelles (csv/tab)  

Fichier.csv : données des points où la distance entre 2 points est > 0, ainsi que des données calculées pour traitements avec un logiciel de type tableur (Excel, LibreOffice, etc)  
Le fichier .csv contient deux champs calculés: "Height" et "VSi" exploitable avec les graphique "Excel 3D-Map"  
- height: hauteur du point par rapport à l'altitude la plus basse du parcours  
- VSi (expérimental):  Vertical Speed Indicator dont les valeurs sont: -1=aucun, 0=faible, 1=moyen, 2=important  
Association conseillée de couleurs au VSi: -1:vert, 0:jaune, 1:rouge, 2:violet  
 
Altitudes corrigées : DEM (Digital Elevation Model) = MNA (modèle numérique d'altitude), de terrain ou de surface  
--- TED (Terrain Elevation Data)   
En raison de l'imprécision des données d'altitudes des appareils enregistrant les données GPS en GPX, pour avoir des altitudes plus précises,  
il est possible d'associer un fichier de données d'altitudes (DEM) en provenance du site www.gpsvisualizer.com/convert  
Pour cela, dans le formulaire de conversion de gpsvisualizer,   
- cocher Output Format = GPX  
- sélectionner "Add DEM elevation data" = ODP1 - Western Europe     (ODP = https:data.opendataportal.at)  
- uploader (convert) le fichier source GPX  
Renommer le fichier obtenu, comme le fichier source, mais en remplacant l'extension .gpx par _ted.gpx (Terrain Elevation Data) 
Puis placer le fichier _ted.gpx dans le même répertoire que le fichier source.
Le fichier obtenu par gpsvisualizer/convert perd les informations de vitesse du fichier gpx source. 
Le script fusionne les données GPX source + données d'altitudes provenant du TED.
Le script vous indiquera la présence du fichier _ted et vous proposera de l'utiliser pour effectuer ses calculs.
Dans ce cas, un nouveau fichier _ted.gpx résultant de la fusion d'avec le gpx source sera alors créé (remplacé).


exemple de statistiques générées:	 
---------------------------------  
Gpx-Filename=20201030_parcours_29km.gpx => nom du fichier gpx   

Gpx-Creator=OsmAnd 3.8.3 => logiciel ayant enregistré la trace dans le fichier gpx  

Gpx-Name=2020-10-25_09-03_Sun => nom interne du fichier gpx  

TrackPoints=719 (moving) / 829 => nbr de points enregistrés dans le fichier (avec le nbr de points où la vitesse > 0)  

AvgRecTimer=10.6 s => fréquence à laquelle le logiciel a enregistré chaque point dans le fichier  

Date=2020-10-25 => date de création du fichier  

StartTime=08:03:01 => heure de début de l'enregistrement  

EndTime=10:29:26 => heure de fin de l'enregistrement  

TotalTime=2 h 26 min => durée du parcours  

MovingTime=2 h 7 min => durée du parcours en mouvement (vitesse>0)  

TotalDistance=28.626 km => distance totale du parcours  

MaxSpeed=31.6 km/h => vitesse maximum  

AvgSpeed=13.5 km/h (moving) => moyenne de vitesse calculée sur les points dont la vitesse > 0 (pour exclure les pauses)  

MinElevation=54.6 m => altitude la plus basse du parcours  

MaxElevation=77.1 m => altitude la plus élevée du parcours  

DiffElevation=22.5 m => différence entre altitude min et altitude max  

TotalAscent=457 m => cumul des montées  

TotalDescent=462 m => cumul descentes  

```  
exemple de contenu de fichier GPX généré par Osmand: (il peut exister des variantes selon l'origine du GPX, en particulier pour le champ "speed" )  
----------------------------------------------------  
<trkpt lat="44.0" lon="-0.5">  
 <ele>18.9</ele>  
 <time>2020-09-29T15:47:54Z</time>  
 <hdop>1.6</hdop>  
 <extensions>  
   <speed>0.324</speed>  
 </extensions>  
</trkpt>  
```
