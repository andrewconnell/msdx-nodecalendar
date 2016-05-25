export interface IEvent {
  id: string;
  startTime: string;
  endTime: string;
  location: string;
  subject: string;
}

export interface IGraphRestEvent {
  'odata.etag': string;
  id: string;
  start?: IGraphRestEventDateTime;
  end?: IGraphRestEventDateTime;
  location?: IGraphRestEventLocation;
  subject?: string;
}

export interface IGraphRestEventDateTime {
  dateTime: string;
  timeZone: string;
}

export interface IGraphRestEventLocation {
  displayName: string;
}
