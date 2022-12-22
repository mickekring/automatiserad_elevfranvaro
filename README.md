# Automatiserad och robotiserad kontroll av elevfrånvaro

Det här är ett script i flera delar som i kort laddar ned en aktuell frånvarorapport på eleverna på Årstaskolan från Stockholms Skolplattform. Därefter läses antalet elever in baserat på klass och sparas i en databas. Ett antalberäkningar senare så spottas lite webbsidor ut som laddas upp på din webbserver. Detta sker typ var 4-5 minut och vips så kan alla ta del av nästan realtidsstatistik kring elevernas frånvaro.

![elevfranvaro](https://user-images.githubusercontent.com/10948066/209207846-3c4c4176-9447-4187-88d8-3d86dbcece59.jpg)
Överbklick av frånvaro med gårdagens siffror inom parentes

![el1](https://user-images.githubusercontent.com/10948066/209207936-4d1d47ac-2ae7-4109-b454-d89c22660faf.jpg)
Graf över senaste 15 dagarna och längre ned hela läsåret

![el2](https://user-images.githubusercontent.com/10948066/209207976-e41514b2-321e-4076-b8c2-13029039c07c.jpg)
Dagens frånvaro baserat på varje klass

## Viktig info

* Det här kommer bara att funka för de som använder Stockholms Skolplattform. Det är en robotisering som härmar en människa som klickar sig fram på webben för att ladda ned rapporten.
* Jag kommer inte att utveckla detta mer, då jag snart inte jobbar i Stockholms stad längre. Ta koden som den är. :)<br />
* Jag kommer heller inte att beskriva så utförligt hur allt fungerar, så du bör ha lite koll på Python om du vill testa. Givetvis så svarar jag på frågor så gott jag kan.

## Installera och köra själv

### Du behöver
* En webbserver med SFTP för uppladdning av frontend
* En dator med Python 3.x som är uppkopplad mot Stockholms stads nät, inloggad med ett riktigt konto

### Installera
* Installera alla moduler som du hittar under import-sektionen längst upp i main.py med pip3 install
* Ladda upp filerna i mappen __upload_to_server__ till webbservern
* Ändra alla inställningar i __config.yml__

## Frågor?
Då är det bara att höra av dig till mig på jag@mickekring.se eller mickekring i de flesta sociala medier

## Läs mer
Du kan läsa lite mer om detta på https://mickekring.se/automatiserad-och-robotiserad-kontroll-av-elevfranvaro/
