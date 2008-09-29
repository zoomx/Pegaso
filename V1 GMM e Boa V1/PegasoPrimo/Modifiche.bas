Attribute VB_Name = "Modifiche"
Option Explicit

'2008 09 29
'Aggiunta una versione alternativa per trovare le porte COM
'Il metodo precedente però elimina le porte COM occupate mentre questo non lo fa.
'Inoltre alcune COM virtuali non sono enumerate in maniera corretta


'2008 09 25
'Corretto baco nel calcolo del CRC
'Il Byte alto del CRC ricevuto era moltiplicato per 255 invece di 256
'Migliorato ma non finito il controllo errori nella ricezione Ymodem


'2008 09 09
'Corretta la routine Fmain bConnect_Click() con l'aggiunta del GoTo in caso di
'mancata risposta del GMM
'Effettuate alcune, ma non tutte, traduzioni in inglese!

'2008 07 16
'In fModem aggiunta la stampa dei timeouts di attesa di risposta del modem remoto
'Cambiata la lunghezza del pacchetto ricevuto da 133 a 132 e da 1029 a 1028 perchè il
'primo byte è catturato a parte

'2008 07 08
'Aggiunto il NAK per il primo pacchetto troppo corto

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

