I think this program is the fast ping in the world.
I surf the Internet and not find the same idea as this ping program.
And try to download many ping programs. But they are not so fast.
In my company, a telecom company, we need to monitor 10 thousands of network devices.
Such as atu-r or router. Use SNMP is so good, but most of these devices do not support SNMP. One day, I find an async ping program from planet source code; I think this must be the solution. So, I start to think how to ping 10 thousands of devices in short time and resolve some problem.
1. Use async ping is good, but not enough, because one async ping application can ping 100 node quick, but there are more than 10 thousand nodes.
2. If I have many async ping agents, may be use multithread is good, but the entire program will limit to the CPU and memory management of OS.
3. In short time, ping 10 thousands of nodes, there are too many messages reply, the User Interface will busy to handle these messages, and then the UI will just like hang.
4. The program might have stable problem.
So, what is my idea?
1. Use async ping as a single ping agent and there are many agents as you want. I try 60 ~ 120 agents are so good in my pc. (I think this is the best idea, right?)
2. A main UI program, load ping list, sends the information to collector and distributes the nodes to ping agents. 
3. Main program send command to ping agent. After ping agent run and get response messages, then it report to the collector program, collector report to main program.
4. UI displays the ping result, collector status and monitors the ping agent¡¦s status.
5. I use access database to store ping list and ping events.
6. All the ping results communication between UI and agents use share memory. (I think this is the best idea too, right?)
7. I know the project can do better, but my company wants to buy expensive software, and not do so good, I mean so fast and so stable program, as my project. So, I stop to continue to modify this program. But this is a good program with many good ideas in it. So, I share it. Hope this can benefit you.
8. My native language is not English, I tried to change some captions and labels to English, hope you understand their meanings.

Install:
1. Main program (the UI), speedping.exe
2. PingAgent.exe and PingCollector.exe must put in ..\agent\ subfolder of the main program.
3. pinglog.db stays with speedping.exe in the same folder.
4. If you want to watch the PingCollector status, change the option to debug mode.

Jimmy Hung,
Email: red.corn@msa.hinet.net

