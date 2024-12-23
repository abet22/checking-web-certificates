from functions import *

if __name__ == "__main__":

    parser = argparse.ArgumentParser(description="Check SSL certificates and save the data into an excel file.")
    parser.add_argument(
        "-e", "--errors", 
        action="store_true", 
        help="Print errors in the terminal instead of saving to misses.txt."
    )
    parser.add_argument(
        "-n", "--namehost",
        type = str,
        help="Check a single host with default port 443."
    )
    parser.add_argument(
        "-p", "--port",
        type = str,
        help="Especify port."
    )
    args = parser.parse_args()

    if args.namehost:
        if args.port:
            tuple_host = (args.namehost, args.port)
            t = get_ssl_info(tuple_host)
            if isinstance(t, dict):
                hostname = t['Hname']
                remaining_days = t['Rdays']
                key_size = t['Ksize']
                date = t['Date']
                print(f"Hostname: {hostname}\nCertificate Expiration Date: {date}\nNumber of remaining days: {remaining_days}\nPublic Key Size: {key_size}")
            else: 
                print (t)
        else: 
            tuple_host = (args.namehost, str(443))
            t = get_ssl_info(tuple_host)
            if isinstance(t, dict): 
                hostname = t['Hname']
                remaining_days = t['Rdays']
                key_size = t['Ksize']
                date = t['Date']
                print(f"Hostname: {hostname}\nCertificate Expiration Date: {date}\nNumber of remaining days: {remaining_days}\nPublic Key Size: {key_size}")
            else: 
                print(t)
    else:
        
        #File with hostnames
        file_path = 'certificate-names.txt'
        
        #Read hostnames
        hostnames = read_hostnames(file_path)

        #Process certificates
        print(f"Checking SSL certificates for {len(hostnames)} domains...\n")
        print(f"Timeout for long response time is {timeout} seconds...\n")
        results = list( process_certificates(hostnames, max_workers=20) )

        save_to_excel(results, args)
