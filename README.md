# Setup

1. Install LibreOffice (https://www.libreoffice.org/download/download/)
2. Run `soffice --accept=socket,host=localhost,port=2002\;urp\;StarOffice.ServiceManager` (the semi-colons are escaped
   here for Linux)
3. Run `./gradlew build`

# Starting LibreOffice in ServerMode

Use `soffice.exe --accept=socket,host=localhost,port=2002;urp;`

# Important Code

`ConnectionAwareClient`

# Useful References

Samples taken from https://api.libreoffice.org/examples/DevelopersGuide/examples.html
Official Documentation: https://api.libreoffice.org/docs/idl/ref/

# Acknowledgements

This project uses utilities written by Andrew Davison
[Java LibreOffice Programming](http://fivedots.coe.psu.ac.th/~ad/jlop/), slightly modified to make it compile against
LibreOffice 6.2.

The utilities are in the `utils` package, and are included as source code instead of a JAR, although we would prefer it
to be published to Maven eventually.
