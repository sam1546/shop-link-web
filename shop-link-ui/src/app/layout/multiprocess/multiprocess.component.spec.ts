import { ComponentFixture, TestBed } from '@angular/core/testing';

import { MultiprocessComponent } from './multiprocess.component';

describe('MultiprocessComponent', () => {
  let component: MultiprocessComponent;
  let fixture: ComponentFixture<MultiprocessComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ MultiprocessComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(MultiprocessComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
