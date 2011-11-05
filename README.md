Node is an Open Source IRC Client for Windows. It is licensed under GPL and developed mainly using VB6, XML, HTML, CSS, and Javascript.

Node was developed mainly in 2004 and 2005. The three ajor contributors are:

 * Dionysis Zindros <dionyziz@gmail.com>
 * Christian Herrmann <happydoener@gmail.com>
 * Josh Nelson <jnfoot@gmail.com>

This repository exists for historical purposes only. See the "Do not use" section for more information.

Some History
============
Node was started by me (dionyziz) and _daemon_ back in 2003. The original idea and name is due to _daemon_, and the project was started as a simple and robust IRC client with few features. It was not meant to be user-friendly but rather, fast and secure. After a month or two, we decided to make the project Open Source and host it on SourceForge.net. As _daemon_ was mostly opposed to that idea, although he consented for code publication, he parted the project a few days after the migration. The project was lead by myself for the next two years. Two other people, ch-world and jnfoot, majorly contributed to the project as developers and administrators, developing many features, making bugfixes, helping it become more popular, and recruiting more people. Several other people, including many developers, translators, beta testers, and skin designers helped out our idea during that time, by developing minor features, translating the program to 9 different languages, and testing the software. We decided to turn it into a very customizable, user-friendly, localized and skinnable client, sacrificing, as expected, speed and memory. The basic idea was that a user not fond of IRC should be able to use it without a huge attempt. Web principles were used, such as point-and-click links for joining channels, opening private windows, and so forth. Many IRC commands, including non-standard, services-based commands were integrated into menus, making it extremely easy for a novice user to go on IRC. Still, security always remained one of our major concerns.

A new version of Node, "Node Lite", was planned in the form of an Instant Messenger for IRC, but it never got far, and no source code from that idea was ever published outside the original team. It was part of an attempt to separate the front-end and the back-end of Node, distinguishing the network-handling and protocol-handling classes from presentation. However, we decided that it wouldn't be a good idea to make it IM-style, because IRC is all about channels, and that an Instant Messaging form is not convenient for its purpose. The project ended up dead, after I parted and, eventually, ch-world as well, in early 2005. Hawkie, another Visual Basic developer, proposed to take over the project, port it to .NET and extend it to become a suite of Internet utilities, including an Internet Explorer-based browser, in 2005. He obtained the NodeIRC.com domain name, but the idea was discontinued after only an unstable and incomplete release.

Innovations
===========
The latest version developed by the original team, 0.35, is very buggy, so even I cannot use it as my Windows IRC Client (I use X-Chat on Windows and irssi on Linux). It contains some nice coding ideas and was one of the largest open source projects in Visual Basic 6 ever developed. (As expected, Visual Basic 6 open source projects are not very popular.) The latest version contains about 31,720 lines of code in Visual Basic 6, XML, HTML, CSS, and Javascript, including the release-shipped skins code. Node achieved memorable activity, including large amounts of work by its developers (we once all worked about 4 hours per day each), notable bug tracking and feature requests activity, and a dedicated userbase. On 18th January 2004, we were the 55th most active project among all SourceForge.net projects, reaching a humble 26 downloads per day for release 0.32 during the first 20 days.

One of the basic ideas that were unique about Node was that, although the basic application was written in Visual Basic, a lot of actual code was written in XML and HTML/CSS/Javascript. XML was used for defining windows structures, and HTML, CSS, and Javascript, were used for rendering everything, including IRC messages. Internet Explorer was used as our rendering engine. This yielded to a wide range of possibilities for what was available for rendering, and made features such as sharing thumbnails on IRC channels available to Node IRC users. It also allowed easy HTML and XML-based skin development and deployment. On the other hand, it severely increased the amounts of memory used by the program and made it less robust. (It also introduced many security and speed issues).

We used NSIS as our installer system.

Things we learned
=================

One of the things we learned through Node development was that interfaces definitions (i.e. GUI specification; a list of controls of a window along with their attributes) should be done using a Content- or Data-describing language (a declarative language) such as HTML or XML, and not through procedural programming. We also learned to localize an application, approaching it using MediaWiki-like ideas, although not as advanced at the time. We discovered that Visual Basic 6 has a lot of bugs, and had to hack the language itself in order to achieve certain abilities without crashing (for example to achieve a "WithEvents" ability on an array of WinSock instances that were not Form ActiveX objects). We had to define our own TCP/IP protocol for exchanging messages between two Node IRC Clients (for abilities such as "Your friend is typing a message" and avatars), and to study the IRC protocol and its services extensions in order to achieve a user-friendly interface that does not require the user to type IRC commands. We also learned to use source version control through CVS, transfer binary data over the Internet securely and other nifty things.

In general, it was lots of fun.

Do not use this
===============
I highly recommend against using this IRC client. It is very buggy and you will have a hard time installing it. It is also dubious that it is compatible with modern versions of the Windows operating system. Furthermore, you may experience many bugs and unstable behavior, including complete and sometimes irrecoverable application crashes. Due to unoptimized code as well a poor choices of libraries, the performance of the software is also dreadful. There are also several security issues I know about that are undocumented and may be harmful to you and your computer (they may, for instance, allow an attacker to send IRC commands on your behalf by utilizing a carefuly crafted XSS attack).

This code is here for two reasons:

 * To embrace it as part of my career and learning path
 * To allow programmers to copy and reuse parts of the code as they see fit

As a developer, you will not find this code very enlightening. The software architecture is not at all sound and violates many well-understood principles of software engineering such as front-end and back-end separation. If you decide to study this code, please understand that it was developed by students who were learning how to program.

License
=======
I am now a supporter of the MIT open source license. However, Node was released under GPL. To show due respect to the decisions made by myself and the team at the time and to honor the fact that I have learned and grown, I am not publishing this software under MIT or BSD, but keeping it as GPL as it was originally intended.

Node IRC - An IRC client for Windows
Copyright (C) 2004 - 2005, Dionysis Zindros and the Node IRC Development Team

This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program; if not, write to the Free Software
Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
