import * as debug from 'debug';
let log: debug.IDebugger = debug('msdx:calendarController');

import {Application, Request, Response} from 'express';
import {User} from '../models/user';
import {AzureAD} from '../auth/azureAD';
import * as config from 'nconf';
import {CalendarService} from '../services/calendarService';
import {IEvent} from '../models/IEvent';
import {CalendarViewModel} from '../models/CalendarViewModel';

export class CalendarController {

  /**
   * @description
   *  Verifies the current user is logged in. If not, they are redirected to the login page.
   * 
   * @param  {Request} req   - Express HTTP Request object.
   * @param  {Response} res  - Express HTTP Response object.
   */
  private static verifyUserLoggedIn(req: Request, res: Response): void {
    let user: User = new User(req);
    if (!user.isAuthenticated()) {
      // redirect to login
      // save path to this page so after logging in, they come back here
      res.redirect('/login?redir=' + encodeURIComponent(req.route.path));
    }
  }

  /**
   * @description
   *  Retrieves an access token for MSFT Graph REST API. 
   *  If this fails, it requests one from Azure AD.
   * 
   * @param  {Request} req   - Express HTTP Request object.
   * @param  {Response} res  - Express HTTP Response object.
   */
  private static getGraphAccessToken(req: Request, res: Response): string {
    // ensure user is logged in
    this.verifyUserLoggedIn(req, res);

    // get the Graph REST API resource ID
    let graphRestApiResourceId: string = config.get('graph-rest-api-resourceid');

    // get the access token to call the endpoint
    let accessToken: string = AzureAD.getAccessToken(req, graphRestApiResourceId);

    // if don't have token / it is expired...
    if (!accessToken || accessToken === 'EXPIRED') {
      // redirect to get token
      // save path to this page so after logging in, they come back here
      res.redirect('/login?' +
        'redir=' + encodeURIComponent(req.route.path) +
        '&resourceId=' + encodeURIComponent(graphRestApiResourceId));
    }

    return accessToken;
  }

  constructor(private app: Application) {
    this.loadRoutes();
  }

  /**
   * Setup routing for controller.
   */
  private loadRoutes(): void {
    log('configuring routes');

    this.app.get('/calendar', this.handleGetCalendar);
  }

  /**
   * @description
   *  Handler for the request for a list of events.
   * 
   * @returns void
   */
  private handleGetCalendar(req: Request, res: Response): void {
    log('handle GET /calendar');

    // get access token for the SPO REST API
    let accessToken: string = CalendarController.getGraphAccessToken(req, res);

    // get missions from services
    let calendarService: CalendarService = new CalendarService(accessToken);
    calendarService.getEvents()
      .then((events: IEvent[]) => {
        log('events received: ' + events.length);

        let vm: CalendarViewModel = new CalendarViewModel(req);
        vm.events = events;

        // render the view
        res.render('calendar/list', vm);
      });
  }
}
