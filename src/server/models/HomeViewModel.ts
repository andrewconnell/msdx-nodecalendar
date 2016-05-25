import {Request} from 'express';
import {BaseViewModel} from './BaseViewModel';

export class HomeViewModel extends BaseViewModel {
  constructor(protected request: Request) {
    super(request);
  }
}
