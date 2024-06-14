# PR-Team-Auto-Allocation
Daily touch-free PR allocation.

Based on the PR allocation made by the PR manager, the PR team works. It took around 30 to 45 minutes per day. Even if the PR manager takes a week off or personal time off, they are still required to manage PR allocation. Once the allocations done, then listings were picked up, the PR team begin efficiently clearing the SLAs. I was instructed to perform some automation on this by my manager.

I took opportunity from the scenario and asked my manager for the necessary access and KT. I started coding the Python code. It's a three-stage process, and I've listed challenges below.
•	The first step was performing a SQL query to get the pending listings from the PR table. Imported necessary packages like "mysql.connector" when using Python. And from my Python window, I got all of the pending listings.
•	The PR monthly roster was used  to find out who was working on the which dates. For that specific need, I utilized the python slice function and obtained the desired outcomes.
•	Using a Python loop, all pending numbers were allocated equally. And Then Python mail function was used to construct and attached the HTML table. There was a new task scheduler created. The time line for this automated allocation is 7 AM every day.
After a one-week test period, the touch-free PR allocation was shared with the PR team starting today.

As a result, automation, I am able to save the manager 30 minutes every day.

