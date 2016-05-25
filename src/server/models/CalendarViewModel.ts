import {Request} from 'express';
import {BaseViewModel} from './BaseViewModel';
import {IEvent} from './IEvent';

export class CalendarViewModel extends BaseViewModel {
  public events: IEvent[];
  public event: IEvent;

  constructor(protected request: Request) {
    super(request);
    this.events = new Array<IEvent>();
  }
}