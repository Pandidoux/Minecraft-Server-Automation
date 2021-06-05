#!/bin/bash
echo "########## Start of Server ##########"
java -jar -Xms256M -Xmx2048M server.jar nogui
echo "########## End of Server ##########"

echo "Exit in 10 seconds..."
sleep 10