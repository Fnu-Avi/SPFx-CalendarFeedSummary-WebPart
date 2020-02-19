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
    id: number|undefined;
}
