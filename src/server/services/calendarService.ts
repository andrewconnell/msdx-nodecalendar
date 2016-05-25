import * as debug from 'debug';
let log: debug.IDebugger = debug('msdx:calendarService');

import * as http from 'http';
import * as config from 'nconf';
import * as request from 'request';
import * as Q from 'q';
import {IEvent, IGraphRestEvent} from '../models/IEvent';

export class CalendarService {

  constructor(private accessToken: string) { }

  /**
   * Retrieves a list of events from the remote data store & returns them within a promise.
   * 
   * @returns Q.Promise<IEvent[]>
   */
  public getEvents(): Q.Promise<IEvent[]> {
    let deferred: Q.Deferred<IEvent[]> = Q.defer<IEvent[]>();

    // create the query
    let queryEndpoint: string = config.get('graph-rest-api-endpoint') +
        '/me/calendar/events' +
        '?$select=start,end,location,subject' +
        '&$filter=isAllDay eq false' +
        '&$orderby=start/dateTime desc';
    log('Graph API query: ' + JSON.stringify(queryEndpoint));

    // create request headers
    let requestHeaders: request.Headers = <request.Headers>{
      'Authorization': 'Bearer ' + this.accessToken,
      'Accept': 'application/json;odata.metadata=minimal'
    };
    log('Graph API request headers: ' + JSON.stringify(requestHeaders));

    // issue request
    request(
      queryEndpoint,
      <request.CoreOptions>{
        headers: requestHeaders,
        method: 'GET'
      },
      (error: any, response: http.IncomingMessage, body: string) => {
        if (error) {
          log('error submitting request to Graph REST API: ' + error);
          deferred.reject(error);
        } else if (response.statusCode !== 200) {
          // if not 200, return error
          let errorMessage: string = '[' + response.statusCode + ']' + response.statusMessage;
          log('error received from Graph REST API: ' + errorMessage);
          deferred.reject(new Error(errorMessage));
        } else {
          log('response body received from Graph REST API: ' + body);

          // collection of events
          let events: IEvent[] = new Array<IEvent>();

          // convert Graph REST data => internal model
          let graphEvents: IGraphRestEvent[] = <IGraphRestEvent[]>(JSON.parse(body)).value;
          for (let graphEvent of graphEvents) {
            events.push(<IEvent>{
              endTime: graphEvent.end.dateTime,
              id: graphEvent.id,
              location: graphEvent.location.displayName,
              startTime: graphEvent.start.dateTime,
              subject: graphEvent.subject
            });
          }

          // return collection of events
          deferred.resolve(events);
        }
      }
    );

    return deferred.promise;
  }
}
