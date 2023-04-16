export class CustomLibraryLibrary {
  public name(): string {
    return 'CustomLibraryLibrary';
  }

  public getCurrentTime(): string {
    let currentDate: Date;
    let str: string;

    currentDate = new Date();

    str = '<br>Todays Date is : ' + currentDate.toDateString();
    str += '<br>Current Time is : ' + currentDate.toTimeString();

    return str;
  }
}
