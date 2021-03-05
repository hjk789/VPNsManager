# VPNs Manager

Easily manage client SoftEther VPN Gate servers, fetch new ones, ping and see max speed

![Main screen](https://i.imgur.com/ono8b9g.png)

## About

This project is a program developed in AutoHotkey to manage client SoftEther VPNs from VPN Gate. Work in progress, alpha stage, there are several bugs, but it's working enough to be useful on a daily basis.

## Key Features

- **Fetch a list of VPN servers from all over the world taken from VPN Gate.**   
By default it fetches 15 servers, with a limit of 5 servers per country, so that your list of servers have a good variety of locations.
- **Hourly port-check each server on the list and know when was the last time it was online.**
- **Register the max speed of the current server so you can know quantitatively how fast it is.**  
It shows both the highest download rate ever reached by the server in one second, and the highest download rate average in 50 samples, which means that servers with a high average are servers with a fast *and* stable connection.
- **Connect to a random server in the list.**
