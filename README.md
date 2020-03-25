# Setup

1. Install LibreOffice (https://www.libreoffice.org/download/download/)
2. Run `soffice --accept=socket,host=localhost,port=2002\;urp\;StarOffice.ServiceManager` (the semi-colons are escaped
   here for Linux)
3. Run `./gradlew build`

# Starting LibreOffice in ServerMode

Use `soffice.exe --accept=socket,host=localhost,port=2002;urp;`

# Useful Stuff

`ConnectionAwareClient`


# Misc

Samples taken from https://api.libreoffice.org/examples/DevelopersGuide/examples.html
