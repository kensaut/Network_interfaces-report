from datetime import datetime
from dotenv import load_dotenv
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
from textfsm.parser import TextFSMError
import time


def get_devices(url, header, netbox=True, devices=None):
    """Gets and returns active devices with IP addresses results from NetBox"""
    requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
    if netbox:
        url_add = "dcim/devices/?status=active&limit=150"
        response = requests.get(
            url=f"{url}{url_add}",
            headers=header, verify=False
        )
        data = response.json()
        devices = data["results"]
        return devices
    else:
        device_list = []
        for device in devices:
            url_add = "ipam/ip-addresses/?address="
            response = requests.get(
                url=f"{url}{url_add}{device}",
                headers=header, verify=False
            )
            data = response.json()
            try:
                devices = data["results"][0]
            except IndexError:
                print(f"{device} is not accounted for in NetBox\n")
            else:
                # pprint(f"INFO: Device result: {device_result}")
                device_list.append(devices)
        return device_list


def get_switches_dictionary(device_result, netbox=True):
    switch_list = []
    if netbox:
        for device in device_result:
            device_name = device["name"]
            device_ip_split = device["primary_ip"]["address"].split("/")
            device_ip = device_ip_split[0]
            switch = {
                "name": device_name,
                "ip": device_ip,
            }
            switch_list.append(switch)
        return switch_list
    else:
        for device in device_result:
            device_name = device["assigned_object"]["device"]["name"]
            device_ip_split = device["address"].split("/")
            device_ip = device_ip_split[0]
            switch = {
                "name": device_name,
                "ip": device_ip,
            }
            switch_list.append(switch)
        return switch_list


def connect_switches(
        switch,
        username,
        password,
        telnet,
        secret,
        show,):
    """Connects to switches and runs show commands"""
    ip = switch["ip"]
    name = switch["name"]
    switch["username"] = username
    switch["password"] = password
    switch["device_type"] = "cisco_ios"
    print("-" * 20)
    print(f"Connecting to {name} - {ip}\n")
    switch.pop("name")
    try:
        connection = ConnectHandler(**switch)
    except netmiko.exceptions.NetmikoAuthenticationException:
        print("Check credentials")
    except KeyboardInterrupt:
        sys.exit()
    except netmiko.exceptions.NetmikoTimeoutException:
        switch["device_type"] = "cisco_ios_telnet"
        try:
            connection = ConnectHandler(**switch)
        except netmiko.exceptions.NetmikoAuthenticationException:
            switch["password"] = telnet
            switch["secret"] = secret
            try:
                connection = ConnectHandler(**switch)
            except EOFError:
                print("Check credentials")
            except netmiko.exceptions.NetmikoAuthenticationException:
                print("Check credentials")
            else:
                # print(f"INFO: Switch prompt: {connection.find_prompt()}")
                # connection.disconnect()
                # report = pull_report(connection, switch, ip, show)
                # return report
                return connection
        else:
            # print(f"INFO: Switch prompt: {connection.find_prompt()}")
            # connection.disconnect()
            # report = pull_report(connection, switch, ip, show)
            # return report
            return connection
    else:
        # print(f"INFO: Switch prompt: {connection.find_prompt()}")
        # connection.disconnect()
        # report = pull_report(connection, switch, ip, show)
        # return report
        return connection

def pull_report(connection, show):
    """Connects to switch and runs show command to gather interfacee details"""
    print("\nGathering interface details.\n")
    print("*" * 25)
    # Gets hostname for printing purposes from show version command
    show_version = connection.send_command(
        "show version", use_textfsm=True,
    )
    hostname = show_version[0]["hostname"]
    try:
        command = connection.send_command(
            show,
            use_textfsm=True,
        )
    except TextFSMError:
        print(f"{hostname} doesn't have {show} available")
    else:
        if show == "show interfaces":
            interface_list = []
            for interface in command:
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
            # pprint(f"INFO: Interfaces list: {interface_list}")
            connection.disconnect()
            report = {
                hostname: interface_list
            }
            # pprint("INFO Switch report: {}".format(switch_report))
            # interface_df = pd.DataFrame(
            #     data=interface_list,
            # )
            return report
        elif show == "show interface switchport":
            switchport_list = []
            for switchport in command:
                switchport_name = switchport["interface"]
                switchport_admin = switchport["admin_mode"]
                switchport_trunking = switchport["trunking_vlans"]
                switchport_details = {
                    "Switchport": switchport_name,
                    "Admin mode": switchport_admin,
                    "Trunking": switchport_trunking
                }
                switchport_list.append(switchport_details)
            connection.disconnect()
            report = {
                hostname: switchport_list
            }
            return report


def write_to_excel(report, path):
    """Takes switch list and writes it to a excel file"""
    name = report.keys()
    df = pd.DataFrame(data=report)
    print(f"\n{df.to_string(index=False)}\n")
    print("*" * 25)
    print(f"\nSaving to {path}\n")
    try:
        df.to_excel(
            f"{path}{name}-interface_{get_date()}.xlsx",
            sheet_name=name,
            index=False,
            na_rep="-".center(1),
        )
    except PermissionError:
        print("File may be open. Please close.\n")
        time.sleep(10)
        try:
            df.to_excel(
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
    # Initiates the environment variables
    load_dotenv()

    # Environment variables to connect to NetBox and authenticate to switches
    NETBOX = os.getenv("NETBOX")
    HEADER = os.getenv("HEADER")
    URL = os.getenv("URL")
    USERNAME = os.getenv("USER")
    PASSWORD = os.getenv("PASSWORD")
    TELNET_PASSWORD = os.getenv("TELNET_PASSWORD")
    SECRET = os.getenv("SECRET")

    # Prepares where to save report
    user = getuser()
    os_system = platform.system()
    if os_system == "Windows":
        inventory_path = f"C:\\Users\\{user}\\Desktop\\"
    else:
        inventory_path = "~/"

    # Arguments that can be passed through to the script
    parser = argparse.ArgumentParser(
        prog="Interface statistics inventory",
        description="Pull a report of interface statistics from a switch",
    )
    parser.add_argument(
        "-u", "--user",
        default=USERNAME,
        help="Username to connect to switches",
    )
    parser.add_argument(
        "-p", "--password",
        nargs="?",
        const="y",
        default=PASSWORD,
        help="Password to connect to switches",
    )
    parser.add_argument(
        "-t", "--telnet",
        nargs="?",
        const="y",
        default=TELNET_PASSWORD,
        help="Telent password if switches connect with telnet",
    )
    parser.add_argument(
        "-s", "--secret",
        nargs="?",
        const="y",
        default=SECRET,
        help="Secret to elevate to priveleged mode",
    )
    parser.add_argument(
        "--show",
        choices=["show interfaces", "show interface switchport"],
        help="Show command to run on switches",
    )
    parser.add_argument(
        "-d", "--devices",
        nargs="*",
        help="List of addresses being checked"
    )
    parser.add_argument(
        "--path",
        default=inventory_path,
        help=f"path to save report: default is {inventory_path}",
    )
    parser.add_argument(
        "--token",
        default=NETBOX,
        help="token to authenticate to NetBox",
    )
    parser.add_argument(
        "--url",
        default=URL,
        help="URL of NetBox server (example: https://10.0.0.5/api/)",
    )
    args = parser.parse_args()
    print(f"INFO: Path: {args.path}")
    device_list = args.devices

    # NetBox's header to authenticate
    header = {
        "Authorization": f"Token {args.token}"
    }

    # Gather credentials to connect to switch
    password = args.password
    telnet = args.telnet
    secret = args.secret
    if password == "y":
        password = getpass("Password: ")
    if args.telnet == "y":
        telnet = getpass("Telnet password: ")
    if args.secret == "y":
        secret = getpass("Secret: ")

    # Checks for device or devices to pull reports from
    if args.devices:
        # print(f"INFO: Devices to test: {args.devices}")
        devices_result = get_devices(
            args.url,
            header,
            netbox=False,
            devices=args.devices
        )
        # pprint(f"INFO: All devices passed through: {devices_result}")
        # pprint(f"INFO: Length of devices: {len(devices_result)}")
        device_list = get_switches_dictionary(devices_result, netbox=False)
        # pprint(f"INFO: Device list not from NetBox:/n {device_list}")
        # print(f"INFO: Length of device list not from NetBox: {len(device_list)}")
    else:
        devices_result = get_devices(args.url, header)
        # pprint(f"INFO: All devices from NetBox: {devices_result}")
        # pprint(f"INFO: Length of devices from NetBox: {len(devices_result)}")
        device_list = get_switches_dictionary(devices_result)
        # pprint(f"INFO: Device list from NetBox: {device_list}")
        # print(f"INFO: Length of device list from NetBox: {len(device_list)}")

    report_collection = []
    for device in device_list:
        # Connects to switches
        connection = connect_switches(
            device,
            args.user,
            password,
            telnet,
            secret,
            args.show,
        )
        # pprint(f"INFO: Connection state: {connection}")
        if connection == None:
            pass
        else:
            report = pull_report(connection, args.show)
            # pprint(f"INFO: Report: {report}")
            if report == None:
                pass
            else:
                if os_system == "Windows":
                    path = args.path + "\\"
                else:
                    path = args.path + "/"
                # write_to_excel(report, path)


if __name__ == "__main__":
    main()
