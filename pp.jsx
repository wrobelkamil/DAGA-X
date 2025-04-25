// Dodajemy listener, który reaguje po otwarciu dokumentu
app.addEventListener("afterOpen", function(event) {
    // Pobranie otwartego dokumentu
    var doc = event.target;
    
    // Sprawdzamy, czy dokument jest zapisany – inaczej nie znamy folderu docelowego
    if (!doc.fullName) {
        $.writeln("Dokument nie jest zapisany. Operacja kopiowania pominięta.");
        return;
    }
    
    try {
        // Iterujemy przez wszystkie łącza w dokumencie
        for (var i = 0; i < doc.links.length; i++) {
            var link = doc.links[i];
            
            // Sprawdzamy, czy nazwa łącza to "przepisy prania.ai"
            if (link.name === "przepisy prania.ai") {
                // Pobieramy obiekt File dla źródłowego pliku łącza
                var sourceFile = new File(link.filePath);
                
                if (sourceFile.exists) {
                    // Określamy folder, w którym znajduje się dokument INDD
                    var destFolder = doc.fullName.parent;
                    
                    // Tworzymy obiekt File dla pliku docelowego (ta sama nazwa)
                    var destFile = new File(destFolder + "/" + sourceFile.name);
                    
                    // Próba skopiowania pliku
                    var copyResult = sourceFile.copy(destFile);
                    
                    if (copyResult) {
                        $.writeln("Plik \"" + sourceFile.name + "\" został skopiowany do folderu: " + destFolder);
                    } else {
                        $.writeln("Nie udało się skopiować pliku: " + sourceFile.fullName);
                    }
                } else {
                    $.writeln("Plik źródłowy nie istnieje: " + sourceFile.fullName);
                }
            }
        }
    } catch(e) {
        $.writeln("Wystąpił błąd podczas przetwarzania dokumentu: " + e);
    }
});
