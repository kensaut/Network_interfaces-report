from datetime import datetime
from dotenv import load_dotenv
from getpass import getpass, getuser
from netmiko import ConnectHandler
from pprint import pprint
from requests.packages.urllib3.exceptions import InsecureRequestWarning
from textfsm.parser import TextFSMError
import argparse
import json
import netmiko.ssh_exception
import openpyxl
import os
import pandas as pd
import platform
import requests
import sys
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


def connect_switches(switch, username, password, telnet, secret, show):
    """Connects to switches and runs show commands"""
    ip = switch["ip"]
    name = switch["name"]
    switch["username"] = username
    switch["password"] = password
    switch["device_type"] = "cisco_ios"
    print("\n" + "*" * 25 + "\n")
    print(f"Connecting to {name} - {ip}\n")
    print("*" * 25 + "\n")
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
                return connection
        else:
            return connection
    else:
        return connection


def pull_report(connection, show):
    """Connects to switch and runs show command to gather interfacee details"""
    # Gets hostname for printing purposes from show version command
    show_version = connection.send_command(
        "show version", use_textfsm=True,
    )
    hostname = show_version[0]["hostname"]
    show_vlan_one = connection.send_command(
        "show interface vlan1", use_textfsm=True,
    )
    ip_split = show_vlan_one[0]["ip_address"].split("/")
    ip = ip_split[0]
    if show == "show interfaces":
        print("Gathering interface details.\n")
        print("*" * 25 + "\n")
        command = connection.send_command(
            show,
            use_textfsm=True,
        )
        report_list = []
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
            report_list.append(interfaces_details)
        disconnect_switch(connection, hostname, ip)
        return report_list, show, hostname
    elif show == "show interface switchport":
        print("Gathering interface details.\n")
        print("*" * 25 + "\n")
        try:
            command = connection.send_command(
                show,
                use_textfsm=True,
            )
        except TextFSMError:
            print(f"{hostname} doesn't have {show} available; using alternative method\n")
            # disconnect_switch(connection, hostname, ip)
        else:
            report_list = []
            for switchport in command:
                switchport_name = switchport["interface"]
                switchport_admin = switchport["admin_mode"]
                switchport_operational = switchport["mode"]
                switchport_access = switchport["access_vlan"]
                switchport_voice = switchport["voice_vlan"]
                switchport_trunking = str(switchport["trunking_vlans"])
                switchport_details = {
                    "Switchport": switchport_name,
                    "Admin mode": switchport_admin,
                    "Operational": switchport_operational,
                    "Access": switchport_access,
                    "Voice": switchport_voice,
                    "Trunking": switchport_trunking
                }
                report_list.append(switchport_details)
            disconnect_switch(connection, hostname, ip)
            return report_list, hostname, ip


def old_switch_switchport(connection, show):
    """Helps with the parsing of individual interfaces"""
    # Gets hostname for printing purposes from show version command
    show_version = connection.send_command(
        "show version", use_textfsm=True,
    )
    hostname = show_version[0]["hostname"]
    show_vlan_one = connection.send_command(
        "show interface vlan1", use_textfsm=True,
    )
    ip_split = show_vlan_one[0]["ip_address"].split("/")
    ip = ip_split[0]
    show_ip_interface_brief = connection.send_command(
        "show ip interface brief",
        use_textfsm=True,
    )
    sliced_show_ip_interface_brief = show_ip_interface_brief[1:]
    # pprint(f"INFO: Sliced show ip int br results: {sliced_show_ip_interface_brief}")
    switchport_report = []
    for interface in sliced_show_ip_interface_brief:
        interface_name = interface["intf"]
        # pprint(f"INFO: Interface name results: {interface_name}")
        show_interface_int_switchport = connection.send_command(
            f"show interface {interface_name} switchport",
            use_textfsm=True,
        )
        # pprint(f"INFO: Switchport results: {show_interface_int_switchport}")
        if "Administrative mode: trunk" in show_interface_int_switchport:
            # print(f"INFO: {interface} is a trunk")
            switchport_admin = "trunk"
        elif "Administrative mode: static access" in show_interface_int_switchport:
            switchport_admin = "static access"
        else:
            # print(f"INFO: {interface} not a trunk")
            switchport_admin = "check interface"
        show_interfaces_int = connection.send_command(
            f"show interfaces {interface_name}",
            use_textfsm=True
        )
        link_status = show_interfaces_int[0]["link_status"]
        protocol_status = show_interfaces_int[0]["protocol_status"]
        description = show_interfaces_int[0]["description"]
        switchport_details = {
            "Switchport": interface_name,
            "Admin mode": switchport_admin,
            "Link status": link_status,
            "Procotol status": protocol_status,
            "Decription": description,
        }
        switchport_report.append(switchport_details)
    disconnect_switch(connection, hostname, ip)
    return switchport_report, hostname, ip


def disconnect_switch(connection, hostname, ip):
    """Disconnects from switch and prints statements"""
    print(f"Disconnecting from {hostname} - {ip}\n")
    print("*" * 25)
    connection.disconnect()


def write_to_excel(report_tuple, command, path, report):
    """Takes switch list and writes it to a excel file"""
    writer = pd.ExcelWriter(f"{path}{command}-{report}_{get_date()}.xlsx")
    for report in report_tuple:
        name = report[1]
        ip = report[2]
        data = report[0]
        df = pd.DataFrame(data=data)
        df.to_excel(
            writer,
            sheet_name=f"{name}_{ip}",
            index=False,
            na_rep="-".center(1),
        )
        print(f"{name}_{ip}\n{df.to_string(index=False)}\n")
        print("*" * 25)
        print(f"\nSaving to {path}\n")
        print("*" * 25)
        writer.save()


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
    URL = os.getenv("URL")

    # Prepares where to save report
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
        "-u", "--user",
        help="Username to connect to switches",
    )
    parser.add_argument(
        "-p", "--password",
        nargs="?",
        const="y",
        help="Password to connect to switches",
    )
    parser.add_argument(
        "-t", "--telnet",
        nargs="?",
        const="y",
        help="Telent password if switches connect with telnet",
    )
    parser.add_argument(
        "-s", "--secret",
        nargs="?",
        const="y",
        help="Secret to elevate to priveleged mode",
    )
    parser.add_argument(
        "-c", "--command",
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
        "--report",
        help="Report name",
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
    device_list = args.devices

    # NetBox's header to authenticate
    header = {
        "Authorization": f"Token {args.token}"
    }

    # Gather credentials to connect to switch
    password = args.password
    telnet = args.telnet
    secret = args.secret
    report_argument = args.report
    if password == "y":
        password = getpass("Password: ")
    if telnet == "y":
        telnet = getpass("Telnet password: ")
    if secret == "y":
        secret = getpass("Secret: ")
    if report_argument == None:
        report_argument = device_list[0]

    # Checks if devices were entered in manually
    if args.devices:
        devices_result = get_devices(
            args.url,
            header,
            netbox=False,
            devices=args.devices
        )
        device_list = get_switches_dictionary(devices_result, netbox=False)
    # No devices entered manually, pulls all devices from NetBox
    else:
        devices_result = get_devices(args.url, header)
        device_list = get_switches_dictionary(devices_result)

    report_collection = []
    for device in device_list:
        # Connects to switches
        ip = device["ip"]
        connection = connect_switches(
            device,
            args.user,
            password,
            telnet,
            secret,
            args.command,
        )
        if connection is None:
            pass
        else:
            report = pull_report(connection, args.command)
            if report is None:
                report = old_switch_switchport(connection, args.command)
                report_collection.append(report)
            else:
                report_collection.append(report)

    # Sets up path, prints report to screen, and creates an Excel spreadsheet
    # of the report
    if report_collection == []:
        pass
    else:
        if os_system == "Windows":
            path = args.path + "\\"
        else:
            path = args.path + "/"
        write_to_excel(report_collection, args.command, path, report_argument)


if __name__ == "__main__":
    main()
