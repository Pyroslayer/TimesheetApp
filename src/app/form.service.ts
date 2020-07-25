import { Injectable } from '@angular/core';
import { FormGroup, FormControl } from '@angular/forms';
import dayjs from 'dayjs';


@Injectable({
  providedIn: 'root'
})
export class FormService {

  constructor() { }

  public assignValueToControl(formControl: FormControl, value: any, validOnly: boolean) {
    const oldValue = formControl.value

    formControl.setValue(value)

    if (validOnly && formControl.invalid) {
      formControl.setValue(oldValue)
    }
  }

  public setToLocalStorage(formGroup: FormGroup, id: string, excludeControls?: string[], maxDaysAge?: number, validOnly?: boolean) {
    let data = {}

    if (maxDaysAge) {
      const validUntil = dayjs().add(maxDaysAge, 'day')
      data['VALID_UNTIL'] = validUntil.toISOString()
    }

    for (let name in formGroup.controls) {
      if (excludeControls && excludeControls.includes(name)) continue

      const control = formGroup.controls[name]
      if (validOnly && control.invalid) continue

      data[name] = control.value
    }

    data = JSON.stringify(data)
    localStorage.setItem(id, <string>data);
  }

  public getFromLocalStorage(formGroup: FormGroup, id: string, excludeControls?: string[], maxDaysAge?: number, validOnly?: boolean) {
    let data = localStorage.getItem(id)
    if (!data) return

    try {
      data = JSON.parse(data)

      // Check if Max-Age is exceeded
      const validUntil = dayjs(data['VALID_UNTIL'])
      if (data['VALID_UNTIL'] && validUntil.isValid() && validUntil.isBefore(dayjs())) {
        localStorage.removeItem(id)
        return
      }

      // Assign items to each accordingcontrol
      for (let name in <any>data) {
        if (!(name in formGroup.controls)) continue
        if (excludeControls && excludeControls.includes(name)) continue

        const control = <FormControl>formGroup.controls[name]
        this.assignValueToControl(control, data[name], validOnly)
      }
    } catch (e) {
      console.log(`Couldn't parse LocalStorage data for formGroup '${id}'`)
      console.log(e)
    }
  }

  public syncWithLocalStorage(formGroup: FormGroup, id: string, excludeControls?: string[], maxDaysAge?: number, validOnly: boolean = true) {
    this.getFromLocalStorage(formGroup, id, excludeControls, maxDaysAge, validOnly)

    formGroup.valueChanges.subscribe(values => {
      this.setToLocalStorage(formGroup, id, excludeControls, maxDaysAge, validOnly)
    })
  }

}