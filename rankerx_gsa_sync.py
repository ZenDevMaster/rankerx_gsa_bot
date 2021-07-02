import logging
import requests
import json
import win32api
import win32gui
import shutil
from requests.exceptions import ConnectionError

logging.basicConfig(filename='log.txt', level=logging.INFO, format='%(asctime)s %(message)s')


with open('config.json') as config_file:
    config = json.load(config_file)
config_file.close()


def saveconfig(config):
    with open("config.json", "w") as config_file:
        write_config = json.dump(config, config_file, indent=4)  # Writing any updates back to he condig file
    print("Config file updated")
    config_file.close()


def get_url_items(id, anchor, s):
    url = config["rankerx_url"] + "/rest/url_items/%s" % id
    response = s.get(url)
    url_list = []
    for u in response.json()["urls"]:
        if anchor:
            url_list.append(u['url']+"#"+u['anchor'])
        else:
            url_list.append(u['url'])
    return url_list


def get_campaign_detail(id, s):
    url = config["rankerx_url"] + "/rest/campaign/%s" % id
    response = s.get(url)
    urls = response.json()
    return urls


def parse_campaign_url_item_ids(campaign_detail):
    url_list_list = []
    for u in campaign_detail["urlLists"]:
        if "Money Site" not in u['name']:
            url_list_list.append(u['id'])
    return url_list_list


def check_is_complete_all_projects(campaign_detail):
    for u in campaign_detail["projects"]:
        if u["status"] != "COMPLETED":
            return False
    return True


def get_campaigns(s):
    url = config["rankerx_url"] + "/rest/campaigns"
    response = s.get(url)
    if response.json()["success"]["message"] == "Success":
        return response.json()
    else:
        return False


def parse_campaign_ids(campaigns):
    campaign_id_list=[]
    for campaign in campaigns["campaigns"]:
        campaign_id_list.append(campaign["id"])
    return campaign_id_list


def is_already_synced(campaign_id):
    if campaign_id in config["campaigns_completed"]:
        return True
    else:
        return False


def get_window_hwnd(title):
    for wnd in enum_windows():
        if title.lower() in win32gui.GetWindowText(wnd).lower():
            return wnd
    return 0


def enum_windows():
    def callback(wnd, data):
        windows.append(wnd)

    windows = []
    win32gui.EnumWindows(callback, None)
    return windows


def refresh_gsa_projects():
    hwnd = get_window_hwnd("GSA Search Engine Ranker")
    if hwnd != 0:
        win32api.SendMessage(hwnd, 1042, 0, 0)
        win32api.SendMessage(hwnd, 1042, 0, 0)  # run it twice to get the Start/Run status back to original state
        print("GSA Project List Refreshed")
    else:
        print("GSA Not Running")


def main():

    logging.info('Started')
    s = requests.session()

    # Authenticate to the RankerX REST api
    post_data = {"username": config["rankerx_username"], "password": config["rankerx_password"], "type": "account"}
    RANKERX_LOGIN = "/public/api/login/account"

    try:
        response = s.post(config["rankerx_url"]+RANKERX_LOGIN, json=post_data)
    except ConnectionError:
        print("Unable to find the RankerX server at %s" % config["rankerx_url"])
        logging.info("Unable to find the RankerX server at %s", config["rankerx_url"])
        quit()

    if response.json()["status"] == "ok":
        print("Successful authentication to RankerX")
    else:
        print("Failed to authenticate to RankerX")
        logging.info("Failed to authenticate to RankerX")
        quit()

    campaigns = get_campaigns(s)
    print("Found %s campaigns" % len(campaigns["campaigns"]))
    for campaign_id in parse_campaign_ids(campaigns):
        campaign_detail = get_campaign_detail(campaign_id, s)
        campaign_name = campaign_detail["campaign"]["name"]

        if check_is_complete_all_projects(campaign_detail) and not is_already_synced(campaign_id):
            print("Working on... %s - %s" % (campaign_id, campaign_name))

            use_rankerx_achor_text = config["use_rankerx_anchor_default"]

            if "[" in campaign_name and "]" in campaign_name:
                # try to grab the project tag
                project_tag = campaign_name[campaign_name.find("[")+len("["):campaign_name.rfind("]")]
                print("Using project tag %s" % project_tag)
                try:
                    project_template_file = config["gsa_project_config"][project_tag]["gsa_prj_template_file"]
                    project_article_file = config["gsa_project_config"][project_tag]["gsa_prj_article_file"]
                    use_rankerx_achor_text = config["gsa_project_config"][project_tag]["use_rankerx_anchor"]
                except KeyError:
                    print("Cannot find tag '%s' in config file. Using default project template" % project_tag)
                    project_template_file = config["gsa_prj_template_file_default"]
                    project_article_file = config["gsa_prj_article_file_default"]
                    pass
            else:
                print("Using default project template")
                project_template_file = config["gsa_prj_template_file_default"]
                project_article_file = config["gsa_prj_article_file_default"]

            all_urls = []
            for list in parse_campaign_url_item_ids(campaign_detail):
                all_urls.extend(get_url_items(list, use_rankerx_achor_text, s))

            logging.info("Discovered a %s URLs in %s", len(all_urls), campaign_name)

            url_string = "{"
            for u in all_urls:
                url_string = url_string + u + "|"
            url_string = url_string[:-1]+"}"

            try:
                project_file = open(project_template_file, "r", encoding="UTF-8")
                project = project_file.read()
                project = project.replace("{{URL_LIST}}", url_string)
                project_file.close()

                filename = config["gsa_prj_path"] + "\\[%s] - %s" % (campaign_id, campaign_name)
                new_project_file = open(filename+".prj", "w", encoding="UTF-8")
                new_project_file.write(project)
                new_project_file.close()

                shutil.copy(project_article_file, filename+".articles")
                print("Successfully configured %s in GSA" % campaign_name)
            except Exception as e:
                print("There was an error while writing files in the GSA project folder.")
                logging.info("There was an error while writing files in the GSA project folder.")
                print(e)

            config["campaigns_completed"].append(campaign_id)

        elif is_already_synced(campaign_id):
            pass
            # already handled by the bot
            # print("Already synced %s" % campaign_id)
        else:
            print("Skipping %s (incomplete tasks remain)" % campaign_detail["campaign"]["name"])

    saveconfig(config)
    refresh_gsa_projects()


if __name__ == '__main__':
    main()
