import { TestBed } from '@angular/core/testing';

import { EmailServiceService } from './email-service.service';

describe('EmailServiceService', () => {
  beforeEach(() => TestBed.configureTestingModule({}));

  it('should be created', () => {
    const service: EmailServiceService = TestBed.get(EmailServiceService);
    expect(service).toBeTruthy();
  });
});
