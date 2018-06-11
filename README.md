# spreadsheetUtil

This restful microservice does the following:

- a request pushes a spreadsheet (xls, xlsx, csv)
- the microservice uses Apache POI to convert the spreadsheet into a JSON object
- a response returns with the JSON object (zipped).

This is a restful web service that uses Spring Boot, Spring Web, as well as Apache POI for all spreadsheet manipulation.

