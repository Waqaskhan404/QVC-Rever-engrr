# AWS Terminal Commands — QVC Bot

---

## Connect to AWS via SSH

```bash
ssh -i "qvt-boot-new.pem" ubuntu@<your-aws-ip>
```

---

## Check If Script Is Running

```bash
ps aux | grep python
```
> Shows all running Python processes. Look for `qvc_book_api.py` in the output.

---

## Check Screen Sessions

```bash
screen -ls
```
> Lists all active screen sessions. A detached session means the bot is running in background.

---

## Start Bot in Screen (Safe Mode)

```bash
screen -S qvc
cd /home/ubuntu/qvc
venv312/bin/python qvc_book_api.py
```
> Starts the booking bot inside a screen session named `qvc`.  
> The bot will keep running even if you close SSH.

**Detach from screen (keep bot running):**
> Press `Ctrl+A` then `D`

---

## Reattach to Running Screen

```bash
screen -r qvc
```
> Reattaches to the `qvc` screen session to view live bot output.

---

## Stop the Bot

```bash
# Find the PID first
ps aux | grep python

# Then kill it using the PID number
kill <PID>
```
> Example: `kill 3468`

---

## Kill Screen Session Completely

```bash
screen -XS qvc quit
```
> Stops and removes the `qvc` screen session entirely.

---

## Pull Latest Code from GitHub

```bash
cd /home/ubuntu/qvc
git pull origin main
```
> Downloads the latest changes from GitHub. Run this before restarting the bot to get the newest version.

---

## Restart Bot (Stop + Relaunch)

```bash
# Step 1: Kill old process
kill <PID>

# Step 2: Pull latest code
cd /home/ubuntu/qvc
git pull origin main

# Step 3: Start new screen session
screen -S qvc
cd /home/ubuntu/qvc
venv312/bin/python qvc_book_api.py
```
> Then detach with `Ctrl+A` then `D`

---

## Check Bot Memory & CPU Usage

```bash
ps aux | grep python
```
> Look at the `%CPU` and `%MEM` columns for `qvc_book_api.py`.

---

## View Live Bot Logs (Inside Screen)

```bash
screen -r qvc
```
> Reattach to see real-time output.  
> Detach again with `Ctrl+A` then `D`

---

## Check How Long Bot Has Been Running

```bash
ps -p <PID> -o pid,etime,cmd
```
> Example: `ps -p 3468 -o pid,etime,cmd`  
> Shows elapsed time since the process started.

---

## Count Saved Captchas

```bash
ls /home/ubuntu/qvc/captcha_solver/real_captchas/ | wc -l
```
> Shows how many captcha images have been collected so far.

---

## Find Project Folder (If Unsure)

```bash
find / -name "qvc_book_api.py" 2>/dev/null
```
> Searches the entire server for the script location.

---

## Quick Reference

| Task | Command |
|---|---|
| Pull latest code | `cd /home/ubuntu/qvc && git pull origin main` |
| Check if running | `ps aux \| grep python` |
| List screen sessions | `screen -ls` |
| Start bot in screen | `screen -S qvc` → `cd /home/ubuntu/qvc` → `venv312/bin/python qvc_book_api.py` |
| Detach from screen | `Ctrl+A` then `D` |
| Reattach to screen | `screen -r qvc` |
| Stop bot | `kill <PID>` |
| Kill screen session | `screen -XS qvc quit` |
| Count captchas | `ls /home/ubuntu/qvc/captcha_solver/real_captchas/ \| wc -l` |
| Push proxies to AWS | `scp -i "...pem" "proxies.txt" ubuntu@13.232.8.193:/home/ubuntu/qvc/` |


---

## Push Proxy Files to AWS

Run these from your **local Windows terminal** (not SSH):

```bash
scp -i "C:\Users\waqas\Desktop\QVC_Production\qvt-boot-new.pem" "C:\Users\waqas\Desktop\QVC_Production\Webshare residential proxies.txt" ubuntu@13.232.8.193:/home/ubuntu/qvc/

scp -i "C:\Users\waqas\Desktop\QVC_Production\qvt-boot-new.pem" "C:\Users\waqas\Desktop\QVC_Production\proxies.txt" ubuntu@13.232.8.193:/home/ubuntu/qvc/
```
> Uploads both proxy files from your PC directly to the AWS server.



