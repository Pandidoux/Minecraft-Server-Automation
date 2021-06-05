@echo off
echo ########## Start of Server ##########
java -jar -Xms256M -Xmx2048M server.jar nogui
echo ########## End of Server ##########


echo Exit in 10 seconds...
ping 127.0.0.1 -n 11 > nul
