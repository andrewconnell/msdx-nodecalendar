import * as debug from 'debug';
let log: debug.IDebugger = debug('msdx:BaseViewModel');
import {Request} from 'express';
import {User} from './user';

export class BaseViewModel {
  public isAuthenticated: boolean = false;
  public userId: string = '';
  public userName: string = '';

  constructor(protected request: Request) {
    this.loadUser();
  }

  /**
   * @description
   *  Loads the currently logged in user (if they are logged in)
   * 
   * @returns void
   */
  private loadUser(): void {
    log('loading currently logged in user / none if not logged in');
    // try to load user
    let user: User = new User(this.request);
    // set user values
    if (user.isAuthenticated) {
      this.isAuthenticated = user.isAuthenticated();
      this.userId = user.userId;
      this.userName = user.fullName();
    }
  }
}
