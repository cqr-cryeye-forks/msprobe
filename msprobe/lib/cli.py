import pathlib

import click
from .exch.exch import *
from .rdp.rdp import *
from .adfs.adfs import *
from .skype.skype import *
from rich.console import Console
from rich.logging import RichHandler
from urllib.parse import urlparse

# Initializing console for rich
console = Console()

# Setting context settings for click
CONTEXT_SETTINGS = dict(help_option_names=["-h", "--help", "help"])


@click.group()
def cli():
    """Find Microsoft Exchange, RD Web, ADFS, and Skype instances"""
    pass


@click.command(no_args_is_help=True, context_settings=CONTEXT_SETTINGS)
@click.option(
    "-v", "--verbose", default=False, required=False, show_default=True, is_flag=True
)
@click.argument("target")
def exch(target, verbose):

    """Process target (if target is URL - delete http(s) schema)"""
    parsed_target = urlparse(target)
    if parsed_target.netloc:
        target = parsed_target.netloc

    """Find Microsoft Exchange servers"""

    # Setting up our console logging
    with console.status("[bold green]Exchange Module Executing...") as status:

        # First trying to find if an Exchange server exists
        exch_endpoint = exch_find(target)

        # Did we find anything
        if exch_endpoint is not None:

            # Checking if OWA and ECP exist
            owa_exists = find_owa(exch_endpoint)
            ecp_exists = find_ecp(exch_endpoint)

            # Getting current Exchange version
            exch_version = find_version(exch_endpoint, owa_exists, ecp_exists)

            # Getting NTLM endpoint information
            exch_ntlm_paths = exch_ntlm_pathfind(exch_endpoint)

            # If no NTLM endpoints found, set data to UNKNOWN, otherwise enumerate
            if len(exch_ntlm_paths) == 0:
                exch_ntlm_info = "UNKNOWN"
            elif len(exch_ntlm_paths) != 0:
                exch_ntlm_info = exch_ntlm_parse(exch_ntlm_paths)

            exch_dict_info = {'URL': exch_endpoint,
                              'VERSION': exch_version,
                              'OWA': owa_exists,
                              'EAC': ecp_exists,
                              'DOMAIN': exch_ntlm_info,
                              'URLS': exch_ntlm_paths}

            exch_dict = {'service': 'Microsoft Exchange', 'data': exch_dict_info}

            status.stop()

            # Displaying info we found
            exch_display(
                exch_endpoint,
                owa_exists,
                ecp_exists,
                exch_version,
                exch_ntlm_paths,
                exch_ntlm_info,
            )

            return exch_dict

        else:
            # exch_dict = {'exch': {'INFO': 'Exchange not found',
            #                       'URL': None,
            #                       'VERSION': None,
            #                       'OWA': None,
            #                       'EAC': None,
            #                       'DOMAIN': None,
            #                       'URLS': [None]
            #                       }}

            """ just empty list? Handle no items in vue"""
            exch_dict = {'service': 'Microsoft Exchange', 'data': None}

            # Logging a failure if no Exchange instance found
            console.log(f"Exchange not found: {target}", style="bold red")
            status.stop()
            return exch_dict


@click.command(no_args_is_help=True, context_settings=CONTEXT_SETTINGS)
@click.option(
    "-v", "--verbose", default=False, required=False, show_default=True, is_flag=True
)
@click.argument("target")
def rdp(target, verbose):

    """Process target (if target is URL - delete http(s) schema)"""
    parsed_target = urlparse(target)
    if parsed_target.netloc:
        target = parsed_target.netloc

    """Find Microsoft RD Web servers"""

    # Setting up our console logging
    with console.status("[bold green]RD Web Module Executing...") as status:

        # First trying to find if an RD Web server exists
        rdpw_endpoint = rdpw_find(target)

        # Did we find anything
        if rdpw_endpoint is not None:

            # Getting the instance version
            rdpw_version = rdpw_find_version(rdpw_endpoint)

            # Getting information about the instance
            rdpw_info = rdpw_get_info(rdpw_endpoint)

            # Getting NTLM endpoint information
            rdpw_ntlm_path = rdpw_ntlm_pathfind(rdpw_endpoint)

            if rdpw_ntlm_path is True:
                rdpw_ntlm_info = rdpw_ntlm_parse(rdpw_endpoint)

            if rdpw_ntlm_path is False:
                rdpw_ntlm_info = [None, None]

            if rdpw_info is not None:
                rdpw_info_list = []
                for i, k in zip(rdpw_info[0::2], rdpw_info[1::2]):
                    rdpw_info_list.append(f"{i} {k}")
            else:
                rdpw_info_list = None

            rdp_dict_info = {'URL': rdpw_endpoint,
                             'VERSION': rdpw_version,
                             'DOMAIN': rdpw_ntlm_info[0],
                             'HOSTNAME': rdpw_ntlm_info[1],
                             'NTLM RPC': rdpw_ntlm_path,
                             'RDPW INFO': rdpw_info_list}

            rdp_dict = {'service': 'Microsoft RD Web', 'data': rdp_dict_info}

            status.stop()

            # Displaying what we found
            rdpw_display(
                rdpw_endpoint,
                rdpw_version,
                rdpw_info,
                rdpw_ntlm_path,
                rdpw_ntlm_info
            )

            return rdp_dict

        else:

            # Logging a failure if no RD Web instance found
            console.log(f"RD Web not found: {target}", style="bold red")

            # rdp_dict = {'rdp': [{'INFO': 'RD Web not found'},
            #                     {'URL': None},
            #                     {'VERSION': None},
            #                     {'DOMAIN': None},
            #                     {'HOSTNAME': None},
            #                     {'NTLM RPC': None},
            #                     {'RDPW INFO': [None]}
            #                     ]}

            rdp_dict = {'service': 'Microsoft RD Web', 'data': None}
            return rdp_dict


@click.command(no_args_is_help=True, context_settings=CONTEXT_SETTINGS)
@click.option(
    "-v", "--verbose", default=False, required=False, show_default=True, is_flag=True
)
@click.argument("target")
def adfs(target, verbose):

    """Process target (if target is URL - delete http(s) schema)"""
    parsed_target = urlparse(target)
    if parsed_target.netloc:
        target = parsed_target.netloc

    """Find Microsoft ADFS servers"""

    # Setting up our console logging
    with console.status("[bold green]ADFS Module Executing...") as status:

        if verbose:
            logging.basicConfig(
                level="DEBUG",
                format="%(message)s",
                handlers=[RichHandler(rich_tracebacks=False, show_time=False)],
            )
            log = logging.getLogger("rich")
            status.stop()
            log.debug("Verbose logging enabled for module: adfs")

        # First trying to find if an ADFS server exists
        adfs_endpoint = adfs_find(target)

        # Did we find anything
        if adfs_endpoint is not None:

            # Getting the instance version
            adfs_version = adfs_find_version(adfs_endpoint)

            # Getting information about ADFS services
            adfs_services = adfs_find_services(adfs_endpoint)

            # Getting information about self-service pw reset endpoint
            adfs_pwreset = find_adfs_pwreset(adfs_endpoint)

            # Getting NTLM endpoint information
            adfs_ntlm_paths = adfs_ntlm_pathfind(adfs_endpoint)
            if len(adfs_ntlm_paths) != 0:
                adfs_ntlm_data = adfs_ntlm_parse(adfs_ntlm_paths)
            else:
                adfs_ntlm_data = "UNKNOWN"


            adfs_dict_info = {'URL': adfs_endpoint,
                              'VERSION': adfs_version,
                              'SSPWR': adfs_pwreset,
                              'URLS': adfs_ntlm_paths,
                              'DOMAIN': adfs_ntlm_data,
                              'SERVICES': adfs_services}

            adfs_dict = {'service': 'Microsoft ADFS', 'data': adfs_dict_info}

            status.stop()

            # Displaying what we found
            adfs_display(
                adfs_endpoint,
                adfs_version,
                adfs_services,
                adfs_pwreset,
                adfs_ntlm_paths,
                adfs_ntlm_data,
            )

            return adfs_dict

        else:

            # Logging a failure if no RD Web instance found
            console.log(f"ADFS not found: {target}", style="bold red")

            # adfs_dict = {'adfs': [{'INFO': 'ADFS not found'},
            #                       {'URL': None},
            #                       {'VERSION': None},
            #                       {'SSPWR': None},
            #                       {'URLS': [None]},
            #                       {'DOMAIN': None},
            #                       {'SERVICES': [None]}
            #                       ]}

            adfs_dict = {'service': 'Microsoft ADFS', 'data': None}
            return adfs_dict


@click.command(no_args_is_help=True, context_settings=CONTEXT_SETTINGS)
@click.option(
    "-v", "--verbose", default=False, required=False, show_default=True, is_flag=True
)
@click.argument("target")
def skype(target, verbose):

    """Process target (if target is URL - delete http(s) schema)"""
    parsed_target = urlparse(target)
    if parsed_target.netloc:
        target = parsed_target.netloc

    """Find Microsoft Skype servers"""

    # Setting up our console logging
    with console.status("[bold green]Skype for Business Module Executing...") as status:

        if verbose:
            logging.basicConfig(
                level="DEBUG",
                format="%(message)s",
                datefmt="[%X]",
                handlers=[RichHandler(rich_tracebacks=False, show_time=False)],
            )
            log = logging.getLogger("rich")
            status.stop()
            log.debug("Verbose logging enabled for module: skype")

        # First trying to find if an SFB server exists
        sfb_endpoint = sfb_find(target)

        # Did we find anything
        if sfb_endpoint is not None:

            # Getting the instance version
            sfb_version = sfb_find_version(sfb_endpoint)

            # Getting information about the instance
            sfb_scheduler = sfb_find_scheduler(sfb_endpoint)
            sfb_chat = sfb_find_chat(sfb_endpoint)

            # Getting NTLM endpoint information
            sfb_ntlm_paths = sfb_ntlm_pathfind(sfb_endpoint)
            if len(sfb_ntlm_paths) != 0:
                sfb_ntlm_data = sfb_ntlm_parse(sfb_ntlm_paths)
            else:
                sfb_ntlm_data = "UNKNOWN"


            if len(sfb_version) > 1:
                sfb_version = f"{sfb_version[0]} ({sfb_version[1]})"

            if sfb_ntlm_data is not None:
                domain_data = sfb_ntlm_data
            elif sfb_ntlm_data is None:
                domain_data = "NOT DEFINED"

            if len(sfb_ntlm_paths) == 0:
                sfb_ntlm_paths = None

            sfb_dict_info = {'URL': sfb_endpoint,
                             'VERSION': sfb_version,
                             'Scheduler': sfb_scheduler,
                             'Chat': sfb_chat,
                             'DOMAIN': domain_data,
                             'URLS': sfb_ntlm_paths}

            sfb_dict = {'service': 'Microsoft Skype', 'data': sfb_dict_info}

            status.stop()

            # Displaying what we found
            sfb_display(
                sfb_endpoint,
                sfb_version,
                sfb_scheduler,
                sfb_chat,
                sfb_ntlm_paths,
                sfb_ntlm_data,
            )

            return sfb_dict

        else:

            # Logging a failure if no SFB instance found
            console.log(
                f"Skype for Business not found: {target}", style="bold red")
            # sfb_dict = {'skype': [{'INFO': 'Skype for Business not found!'},
            #                       {'URL': None},
            #                       {'VERSION': None},
            #                       {'Scheduler': None},
            #                       {'Chat': None},
            #                       {'DOMAIN': None},
            #                       {'URLS': [None]}
            #                       ]}

            sfb_dict = {'service': 'Microsoft Skype', 'data': None}
            return sfb_dict


@click.command(no_args_is_help=True, context_settings=CONTEXT_SETTINGS)
@click.option(
    "-v", "--verbose", default=False, required=False, show_default=True, is_flag=True
)
@click.argument("target")
@click.pass_context
def full(ctx, target, verbose):
    """Find all Microsoft supported by msprobe"""

    final_results_list = []

    data_exch = ctx.forward(exch)
    final_results_list.append(data_exch)

    data_rdp = ctx.forward(rdp)
    final_results_list.append(data_rdp)

    data_adfs = ctx.forward(adfs)
    final_results_list.append(data_adfs)

    data_skype = ctx.forward(skype)
    final_results_list.append(data_skype)

    root_path = pathlib.Path(__file__).parent.parent.parent
    result_file_path = root_path.joinpath("result.json")
    result_file_path.write_text(json.dumps(final_results_list))


# Defining commands
cli.add_command(exch)
cli.add_command(adfs)
cli.add_command(skype)
cli.add_command(rdp)
cli.add_command(full)

if __name__ == "__main__":
    cli()
