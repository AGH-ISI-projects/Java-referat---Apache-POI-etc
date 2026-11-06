# CSV do Excel - Konwerter

Prosty projekt edukacyjny w Javie prezentujący wykorzystanie i integrację dwóch bibliotek:
* **Apache Commons CSV** – do bezpiecznego odczytu i parsowania płaskich plików `.csv`.
* **Apache POI** – do generowania i zapisu danych jako arkusz kalkulacyjny `.xlsx`.

Program wczytuje testową listę z pliku `dane.csv` i natychmiast przekształca ją w gotowy plik `wynik.xlsx`.

## Wymagania
* Java 17+ (lub nowsza)
* Maven

## Jak uruchomić

Projekt jest skonfigurowany tak, aby budować gotowy do uruchomienia plik JAR ze wszystkimi zależnościami (Fat JAR).

1. Zbuduj projekt:
```bash
mvn clean package
```

2. Uruchom aplikację:
```bash
java -jar target/przyklad-1.0-SNAPSHOT-jar-with-dependencies.jar
```

**Ważne:** Upewnij się, że plik wejściowy `dane.csv` znajduje się w głównym katalogu projektu (tam gdzie plik `pom.xml`) przed uruchomieniem aplikacji.