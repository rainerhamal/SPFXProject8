//! Lsn 4.3.4 Update the public interface for the React component. This is the interface for the public properties of the React component. It will need to display a list of items, so update the signature to accept a list of items.
import {
  ButtonClickedCallback,
  ICountryListItem
} from '../../../models';

//! Lsn 4.3.5 Update the interface to replace the existing description property to be a collection of items to be passed in and add an event when a button is selected:
export interface ISpfxHttpClientDemoProps {
  // description: string;

  spListItems: ICountryListItem[];
  onGetListItems?: ButtonClickedCallback;

  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
