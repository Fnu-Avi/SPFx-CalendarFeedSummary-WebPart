/**
 * ExtensionService
 */
import { HttpClientResponse } from "@microsoft/sp-http";
import { ICalendarService } from "..";
import { BaseCalendarService } from "../BaseCalendarService";
import { ICalendarEvent } from "../ICalendarEvent";
import { Web } from "@pnp/sp";
import { combine } from "@pnp/common";

export class SharePointCalendarService extends BaseCalendarService
  implements ICalendarService {
  constructor() {
    super();
    this.Name = "SharePoint";
  }

  public getEvents = async (): Promise<ICalendarEvent[]> => {
    const parameterizedFeedUrl: string = this.replaceTokens(
      this.FeedUrl,
      this.EventRange
    );

    // Get the URL
    let webUrl = parameterizedFeedUrl.toLowerCase();

    // Break the URL into parts
    let urlParts = webUrl.split("/");

    // Get the web root
    let webRoot = urlParts[0] + "/" + urlParts[1] + "/" + urlParts[2];

    // Get the list URL
    let listUrl = webUrl.substring(webRoot.length);

    // Find the "lists" portion of the URL to get the site URL
    let webLocation = listUrl.substr(0, listUrl.indexOf("lists/"));
    let siteUrl = webRoot + webLocation;

    // Open the web associated to the site
    let web = new Web(siteUrl);

    // Get the web
    await web.get();
    // Build a filter so that we don't retrieve every single thing unless necesssary
    let dateFilter: string = "EventDate ge datetime'" + this.EventRange.Start.toISOString() + "' and EndDate lt datetime'" + this.EventRange.End.toISOString() + "' and EventStatus eq 'Approved'";
    try {
      const items = await web.getList(listUrl)
        .items
        .select("Id,Title,Description,EventDate,EndDate,fAllDayEvent,Category,Location, EventStatus")
        .orderBy('EventDate', true)
        .filter(dateFilter)
        .get();
      // console.log(web.getList(listUrl).items.get());
      // console.log(items);
      // Once we get the list, convert to calendar events
      // let el = document.createElement( 'html' );
      // var parser = new DOMParser();
      let events: ICalendarEvent[] = items.map((item: any) => {
        let eventUrl: string = combine(webUrl, "DispForm.aspx?ID=" + item.Id);
        const eventItem: ICalendarEvent = {
          title: item.Title,
          start: item.EventDate,
          end: item.EndDate,
          url: eventUrl,
          allDay: item.fAllDayEvent,
          category: item.Category,
          // description: item.Description,
          description: "<body id='pi'>"+item.Description+"</body>",
          location: item.Location,
          approvalStatus: item.EventStatus,
          eventCampus: item.EventCampus,
          id: item.Id
        };
        
        
        // var htmlDoc = parser.parseFromString(eventItem.description, 'text/xml');
        // console.log(htmlDoc);
        // console.log(document.createElement( 'html' ).innerHTML = "<html>"+eventItem.description+"</html>");
        // console.log(el);
        // console.log(eventItem.description);
        return eventItem;
      });
      // console.log(events);
      // Return the calendar items
      return events;
    }
    catch (error) {
      console.log("Exception caught by catch in SharePoint provider", error);
      throw error;
    }
  }
}
