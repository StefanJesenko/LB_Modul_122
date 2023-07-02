# LB_Modul_122 Portfolio

Erstell von Stefan Jesenko

## Einleitung

In meinem Projekt habe ich ein Skript erstellt, dieses kann die PowerPoint und Word Dokumente in einem Ordner auf nichtmehr funktionierende Links überprüfen.

## Was habe ich gelernt?

Ich habe gelernt wie man mit PowerShell einen Link überprüfen kann.

## Beschreibung

Mit PowerShell kann man mit dem begriff `HyperLinks` auf Links zugreifen. In meinem Skript öffne ich zuerst ein Dokument dann wird jeder Link in dem Dokument getestet. Das mache ich mit einer `foreach` Schleife, in der wird jeder Link geöffnet und dann wird getestet ob der Link funktioniert mit `Test-Connection`.

```PS1
foreach ($file in $pptFiles) {
        $presentation = $pptApp.Presentations.Open($file.FullName)
        
        foreach ($slide in $presentation.Slides) {
            foreach ($link in $slide.Hyperlinks) {
                $linkAddress = $link.Address
                $response = Test-Connection $linkAddress -Count 1 -Quiet
                if ($response) {
                
                }
                else {
                    Write-Host "Link is not accessible in PowerPoint document $($file.Name): $linkAddress"
                }
            }
        }
        
        $presentation.Close()
        Write-Host "Closed PowerPoint presentation: $($file.Name)"
        Write-Host "---------------------------------------------"
    }
```

Hier sieht man wie ich die `Hyperlinks` in der `foreach` auf funktionalität getestet habe.

## Reflexion und Fazit

Bei mir ist gut gelaufen:
Ich hatte bei meinem Skript direkt eine Idee wie ich dieses Skript mit Funktionen Umsetzen kann.

Bei mir ist nicht gut gelaufen:

Ich hatte probleme die Ordner in dem Hauptordner zu kontrolieren.

## Fazit
