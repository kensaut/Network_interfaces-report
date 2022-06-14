from datetime import datetime
from getpass import getpass, getuser
from netmiko import ConnectHandler
from pprint import pprint
from requests.packages.urllib3.exceptions import InsecureRequestWarning
import argparse
import config
import netmiko.ssh_exception
import openpyxl
import os
import pandas as pd
import platform
import requests
import sys
import time


NETBOX = config.NETBOX
URL = config.URL
INVENTORY_PATH = config.INVENTORY_PATH


def get_device(ip, header, url):
    """
        Gets IP information passed into function and returns results from
        NetBox
    """
    url_add = "ipam/ip-addresses/?address="
    requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
    response = requests.get(
        url=f"{url}{url_add}{ip}",
        headers=header,
        verify=False,
    )
    device = response.json()
    results = device["results"][0]
    switch_name = results["assigned_object"]["device"]["name"]
    return switch_name


def get_device_details(switch_name, header, url):
    """
        Gets device information passed into function and returns results from
        NetBox
    """
    url_add = "dcim/devices/?name="
    requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
    response = requests.get(
        url=f"{url}{url_add}{switch_name}",
        headers=header,
        verify=False,
    )
    device = response.json()
    results = device["results"][0]
    return results


def get_credentials(connection):
    """Takes connection preference and returns credentials in a dictionary"""
    if connection == "telnet":
        password = getpass("Telnet password: ")
        credentials = {
            "password": password,
        }
        return credentials
    if connection == "ssh":
        username = input("Username: ")
        password = getpass("Password: ")
        credentials = {
            "username": username,
            "password": password,
        }
        return credentials


def switch_prep(results, connection, credentials):
    """Takes devices and credentials and creates a dictionary of switches for
    connection"""
    ip = results["primary_ip"]["address"].removesuffix("/16")
    name = results["name"]
    switch = {
        "name": name,
        "ip": ip,
        "device_type": connection,
    }
    if connection == "cisco_ios":
        switch["username"] = credentials["username"]
        switch["password"] = credentials["password"]
    else:
        switch["password"] = credentials["password"]
    name = results["name"]
    return switch


def connect_switch(switch):
    """Connects to switch and runs show command to gather interfacee details"""
    ip = switch["ip"]
    name = switch["name"]
    switch.pop("name")
    print(f"\nConnecting to {name} - {ip}\n")
    print("*" * 25)
    try:
        connection = ConnectHandler(**switch)
    except netmiko.exceptions.NetmikoAuthenticationException:
        print("\nAccess denied or wrong credentials.")
        sys.exit()
    except netmiko.exceptions.NetmikoTimeoutException:
        print("\nCheck connection type.")
        sys.exit()
    except KeyboardInterrupt:
        sys.exit()
    else:
        print("\nGathering interface details.\n")
        print("*" * 25)
        interfaces = connection.send_command(
            f"show interface | exclude Vlan1",
            use_textfsm=True,
        )
        interface_list = []
        for interface in interfaces:
            interface_name = interface["interface"]
            interface_description = interface["description"]
            if interface_description == "":
                interface_description = "NO DESCRIPTION"
            interface_bandwidth = interface["bandwidth"]
            interface_link = interface["link_status"]
            interface_protocol = interface["protocol_status"]
            interfaces_details = {
                "Interface": interface_name,
                "Int description": interface_description,
                "Int bandwidth": interface_bandwidth,
                "Int link status": interface_link,
                "Int protocol status": interface_protocol,
            }
            interface_list.append(interfaces_details)
        connection.disconnect()
        interface_df = pd.DataFrame(
            data=interface_list,
        )
        return interface_df


def write_to_excel(report, switch, path, name):
    """Takes switch list and writes it to a csv file"""
    print(f"\n{report.to_string(index=False)}\n")
    print("*" * 25)
    print(f"\nSaving to {path}\n")
    try:
        report.to_excel(
            f"{path}{name}-interface_{get_date()}.xlsx",
            sheet_name=name,
            index=False,
            na_rep="-".center(1),
        )
    except PermissionError:
        print("File may be open. Please close.\n")
        time.sleep(10)
        try:
            report.to_excel(
                f"{path}{name}-interface_{get_date()}.xlsx",
                sheet_name=name,
                index=False,
                na_rep="-".center(1),
            )
        except PermissionError:
            print("File was open during program. Ending program.")
            sys.exit()


def get_date():
    """Returns current timestamp in YYYY-MM-DD"""
    now = datetime.now()
    timestamp = now.strftime("%Y-%m-%d")
    return timestamp


def main():
    user = getuser()
    os_system = platform.system()
    if os_system == "Windows":
        inventory_path = f"C:\\Users\\{user}\\Desktop\\"
    else:
        inventory_path = "~/"
    parser = argparse.ArgumentParser(
        prog="Interface statistics inventory",
        description="Pull a report of interface statistics from a switch",
    )
    parser.add_argument(
        "devices",
        help="List of addresses being checked"
    )
    parser.add_argument(
        "-c", "--connection",
        default="ssh",
        help="""
            Protocol used to connect to device, SSH, Telnet, etc. default is ssh
        """,
    )
    parser.add_argument(
        "-t", "--token",
        default=NETBOX,
        help="token to authenticate to NetBox",
    )
    parser.add_argument(
        "-u", "--url",
        default=URL,
        help="URL of NetBox server (example: https://10.0.0.5/api/)",
    )
    parser.add_argument(
        "-p", "--path",
        default=inventory_path,
        help=f"path to save report: default is {inventory_path}",
    )
    args = parser.parse_args()
    device_list = args.devices
    header = {
        "Authorization": f"Token {args.token}"
    }
    device_result = get_device(device_list, header, args.url)
    device_details = get_device_details(device_result, header, args.url)
    switch_name = device_details["name"]
    credentials = get_credentials(args.connection.lower())
    if args.connection.lower() == "telnet":
        device_connection = "cisco_ios_telnet"
    else:
        device_connection = "cisco_ios"
    switch_dictionary = switch_prep(device_details, device_connection, credentials)
    interface_report = connect_switch(switch_dictionary)
    if os_system == "Windows":
        path = args.path + "\\"
    else:
        path = args.path + "/"
    write_to_excel(interface_report, switch_dictionary, path, switch_name)


if __name__ == "__main__":
    main()
