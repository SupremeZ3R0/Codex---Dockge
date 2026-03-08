# Codex---Dockge

Dockge stacks via OpenAI Codex.

## TrueNAS SCALE 25.10.1 (Goldeye): Installing ArchiSteamFarm in Dockge

This repo now documents a practical setup based on your layout:

- TrueNAS SCALE: `25.10.1 (Goldeye)`
- `apps` dataset (SSD): app installs
- `configs` dataset (SSD): persistent app configs
- `stacks` dataset (SSD): Dockge stacks storage
- Dockge installed from TrueNAS App Store with **Dockge Stacks Storage** pointed at `stacks`

---

### 1) Verify where Dockge writes stack projects

In Dockge, create a test stack and confirm it creates a folder under your stacks dataset, typically something like:

- `/mnt/<pool>/stacks/<stack-name>/compose.yaml`

If this works, your Dockge "Stacks Storage" mapping is correct.

---

### 2) Prepare persistent ASF directories on TrueNAS

Create dedicated paths for ASF so updates/redeploys keep bot state:

- `/mnt/<pool>/configs/archisteamfarm/config`
- `/mnt/<pool>/configs/archisteamfarm/plugins` (optional)
- `/mnt/<pool>/configs/archisteamfarm/logs` (optional)

You can create these in **Datasets** (nested datasets) or as directories inside `configs`.

---

### 3) Create the ASF stack in Dockge

In Dockge, create a new stack named `archisteamfarm` and use:

```yaml
services:
  asf:
    image: justarchi/archisteamfarm:latest
    container_name: archisteamfarm
    restart: unless-stopped
    environment:
      - TZ=UTC
    volumes:
      - /mnt/<pool>/configs/archisteamfarm/config:/app/config
      - /mnt/<pool>/configs/archisteamfarm/plugins:/app/plugins
      - /mnt/<pool>/configs/archisteamfarm/logs:/app/logs
    ports:
      - "1242:1242"   # ASF IPC web UI/API
```

Then deploy the stack.

Notes:

- Keep paths as **absolute** `/mnt/...` TrueNAS paths.
- If you don't need plugins/log bind mounts, you can remove those two lines.

---

### 4) First-time ASF bot config

After first deploy, place your bot JSON in:

- `/mnt/<pool>/configs/archisteamfarm/config/<BotName>.json`

And usually global settings in:

- `/mnt/<pool>/configs/archisteamfarm/config/ASF.json`

Then restart the container from Dockge.

---

### 5) Permissions checks (important on TrueNAS)

If ASF can't read/write config files:

1. Ensure dataset ACL/permissions permit the app/container user to read/write.
2. For quick diagnostics, temporarily grant permissive rights on the ASF config path, test startup, then tighten ACLs.
3. Confirm no stale root-owned files block writes.

A common symptom is ASF starting with missing bot config or repeatedly regenerating defaults.

---

### 6) Networking and access

- IPC/UI is on port `1242` if enabled by your ASF config.
- If you already run reverse proxy, you can remove host port exposure and proxy internally.
- Keep ASF off host networking unless you specifically need it.

---

### 7) Upgrade workflow

With Dockge, upgrades are straightforward:

1. Pull latest image (Dockge update/redeploy action).
2. Recreate container.
3. Persistent state remains in `/mnt/<pool>/configs/archisteamfarm/...`.

---

## Troubleshooting quick list

- **Stack deploys but ASF exits immediately**: check container logs in Dockge; usually config/permission issue.
- **No bot appears**: verify `<BotName>.json` location and JSON validity.
- **IPC unreachable**: confirm port mapping and ASF IPC settings in `ASF.json`.
- **Data lost after redeploy**: ensure volumes point to `/mnt/<pool>/configs/...`, not container-only paths.

