Attribute VB_Name = "Modifiche"
Option Explicit

'2008 05 13
'Commentate le parti di comunicazione YMODEM che impediscono la compilazione

'2008 05 19
'Aggiunta la traduzione del file PTM tecnico

'2008 05 20
'Corretta formula WeekDayn = DateMinsMet / 1440 + 1
'in WeekDayn = Int(DateMinsMet / 1440) + 1 che dava errore di arrotondamento
'Sostituito Hour con Hours (Hour è il nome di una funzione)

'2008 05 21
'Aggiunta una label di stato
'Viene aggiornata dopo la connessione del modem
'Aggiunte le operazioni a bit shift left e right nell'apposito modulo
'Adesso la casella di testo mostra cosa sta accadendo
'Reimplementazione dell'YMODEM ancora non completa
'Il sorgente si compila senza errori

'2008 05 27
'Reimplementazione dell'YMODEM da zero. Sembra funzioni
'Inizio implementazione chiamata modem e connessione a GMM

'2008 05 28
'Cambiata l'estensione predefinita dei file tradotti da .dat a .csv
'Fatte alcune modifiche all'interrogazione del GMM

