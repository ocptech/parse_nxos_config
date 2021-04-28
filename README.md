# Parse NXOS Config :white_check_mark:
[![published](https://static.production.devnetcloud.com/codeexchange/assets/images/devnet-published.svg)](https://developer.cisco.com/codeexchange/github/repo/pablog86/aci-contractchecker)

## Description

This example script generates an Excel file with the information gathered from running-config file from Cisco NXOS.

The file contains the following pages:
- VLANs: Table with vlan information.
- SVIs: Table with vlan interfaces information.
- Ints: Table with interface information.
- Po: Table with port-channel information.


## Requirements

Running config file in TXT format.

### Clone the repository

```text
git clone https://github.com/ocptech/parse_nxos_config
cd parse_nxos_config

chmod 755 parse_nxos_config
```

### Python environment

Create virtual environment and activate it (optional)

```text
python3 -m venv parse_nxos_config
source parse_nxos_config/bin/activate
Install required modules
```

Install required modules

```text
pip install -r requirements.txt
```


## Usage & examples

Just run the parse-conf.py script and select the config file. The output will be generated in the same directory.


