import { ChoiceFieldFormatType } from "@pnp/sp";

export interface ICalendarEvent {
    title: string;
    start: Date;
    end: Date;
    url: string|undefined;
    allDay: boolean;
    category: string|undefined;
    description: string|undefined;
    location: string|undefined;
    approvalStatus: string|undefined;
    eventCampus: ChoiceFieldFormatType|undefined;
    id: number|undefined;
}
