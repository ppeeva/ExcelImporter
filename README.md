# ExcelImporter

Example piece of functionality - WEB API method for importing Excel files. Tailored for custom data models. Uses unsupported version of Epplus. Returns json echoing the file data. Uses reflection. Option to export Excel file template.

# Technologies

C# (ASP .Net Core Web API)

# Postman requests

Method: POST

URL: https://localhost:44350/api/import/testFile

Body: form-data

Body: Content - File


----


Method: GET

URL: https://localhost:44350/api/import/exportFileTemplate
