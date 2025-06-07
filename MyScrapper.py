import asyncio
import re
import time
from datetime import datetime
from typing import Annotated, Literal

import aiohttp
from bs4 import BeautifulSoup, Tag
from icecream import ic

BASE_URL = "https://wol.jw.org/"


async def extract_time_duration(element: str):
    match = re.search(r"\((\d+ min\.)\)", element)
    return match.group(0) if match else None


NUMBERED_ENTRY: re.Pattern = re.compile(r"^(?P<number>\d+)\.\s(?P<value>.*)$")


async def get_time_excerpt(element: Tag):
    text = element.text.strip()

    time_duration = await extract_time_duration(
        element.find_next_sibling().find("p").text.strip()
    ) if element.find_next_sibling() is not None else ""

    m = NUMBERED_ENTRY.match(text)
    if not m:
        return text
    number = int(m['number'])
    value = m['value']
    return number, value + " " + time_duration


async def _parse_url(url: str):
    async with aiohttp.ClientSession() as session:
        async with session.get(url) as response:
            html = await response.text()
            return BeautifulSoup(html, "html.parser")


async def _extract_weeks_and_content(soup: BeautifulSoup) -> tuple[
    Annotated[str, "Cover Page Title"],
    Annotated[list, "List containing dates of each of the weeks"],
    Annotated[list, "List containing the links of each of the weeks"]
]:
    """
    Takes the soup object parsed from the url and extracts the number of weeks as well as the titles of each week.

    :param soup:

    :return: str - Cover title
    :return: list - List of dates for each week
    :return: list - List of links for each week
    """

    weeks = soup.find_all(class_=["row", "card"])  # Finding all the elements that contain the week content
    no_of_weeks = len(weeks) - 1 # Excluding the cover page gotten from the url
    print("Number of weeks in the month: ", no_of_weeks, "\n\n")

    cover_page = weeks[0]  # The cover page
    cover_title: str = (
        cover_page
        .find(class_=["cardTitleBlock"])
        .text
        .strip()
    )  # Extracting the text for each section (useful for getting things like title and week dates)
    print(f"Document for {cover_title} is being processed. Please be patient. ")
    date_list = []
    links_per_page = []

    weeks = weeks[1:]  # The rest of the weeks' content
    for week in weeks:
        links_per_page.append(week.a["href"])  # storing all the lists of the weeks' content
        date_list.append(
            week
            .a
            .find(class_="cardTitleBlock")
            .text
            .strip()
        )  # storing all the dates/times of the weeks.

    return cover_title, date_list, links_per_page


async def _parse_week_content(url: str):
    print(f"Attempting to parse url {url} ...")
    parse_dict = {}

    soup = await _parse_url(BASE_URL + url)

    print("\nGotten contents of URL..Parsing ")
    start_time = time.perf_counter()

    try:
        date = soup.find("header").h1.text
    except AttributeError:
        print("ERROR FINDING `header` - SKIPPING... ")
        return None

    pattern_for_song = re.compile(r'\bSong\s+\d+')

    try:
        meeting_content = [
            await get_time_excerpt(element)
            for element in soup.find(class_="bodyTxt").find_all(class_="du-fontSize--base")
            if not pattern_for_song.search(element.text.strip())
        ]
    except AttributeError:
        print("ERROR FINDING `header` - SKIPPING... ")
        return None

    parse_dict["date"] = date
    parse_dict["main_table"] = meeting_content

    print(f"Done with week of {date}...")
    end_time = time.perf_counter()
    print(f"finished in {end_time - start_time:.4f} seconds")
    return parse_dict


async def parse_workbook_url(month: Literal["January", "March", "May", "July", "September", "November"]):
    print("Getting contents of url ...")
    soup = await _parse_url(
        f"https://wol.jw.org/en/wol/library/r1/lp-e/all-publications/meeting-workbooks/life-and-ministry-meeting-workbook-{str(datetime.now().year)}/" + month.lower())

    cover_title, dates, links_per_page = await _extract_weeks_and_content(soup)

    async_tasks = []  # Gather all web scrapping tasks and make them run at the same time to make them run faster.

    for link in links_per_page:
        async_tasks.append(_parse_week_content(link))
    results = await asyncio.gather(*async_tasks)
    return cover_title, results


async def main():
    for result in await parse_workbook_url("September"):
        ic(result)

if __name__ == "__main__":
    asyncio.run(main())
