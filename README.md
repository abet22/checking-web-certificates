This Python program allows you to check SSL/TLS certificates for multiple hosts. You can check all certificates at once or check a specific one. If any errors occur during the certificate check, these errors will be saved in a file called `misses.txt`.

## Features

The program has several execution options:

- **Mode to check all certificates** for the hosts listed in the `certificate-names.txt` file, some of which may have special ports.
- **Mode to check a single certificate** given a hostname, without specifying a port.
- **Mode to check a single certificate** given a hostname, specifying a port.
- **Mode to show errors in the terminal** instead of saving them to the `misses.txt` file.

## Requirements

- Python 3.x
- A .txt file with all hostsnames you want to check

## Usage

You can get more details on how to use the program by running the following command in your terminal:
- python3 check-certificates.py -h

Here are some example executions:

- Check all certificates for hosts in certificate-names.txt:
    python3 check-certificates.py

- Check a single host (without specifying a port):
    python3 check-certificates.py -n {hostname}

- Check a single host with a specified port:
    python3 check-certificates.py -n {hostname} -p {port}

- Print errors in the terminal instead of saving to misses.txt:
    python3 check-certificates.py -e

## Generated Files

- misses.txt: A file where any errors encountered during the certificate check are logged.
- certificates_"%Y-%m-%d".xlsx: An excel file with the important information of each certificate.
