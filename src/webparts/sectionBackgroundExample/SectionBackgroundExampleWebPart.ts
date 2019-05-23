import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { ISemanticColors, ThemeChangedEventArgs, ThemeProvider, IReadonlyTheme } from '@microsoft/sp-component-base';
import { getTheme, ITheme } from '@microsoft/office-ui-fabric-react-bundle';

export interface ISectionBackgroundExampleWebPartProps {
  description: string;
}

export default class SectionBackgroundExampleWebPart extends BaseClientSideWebPart<ISectionBackgroundExampleWebPartProps> {
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme;

  // 注册一个 themeChangedEvent
  // 当页面的 theme 发生变化时，这个event就会被触发
  protected onInit(): Promise<void> {
    // get the page theme
    const pageTheme: ITheme = getTheme();
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this._themeVariant = this._themeProvider.tryGetTheme();

    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);

    return super.onInit();
  }

  // 移除 event listener 来防止内存泄漏 
  protected onDispose(): void {
    this._themeProvider.themeChangedEvent.remove(this, this._handleThemeChangedEvent);

    super.onDispose();
  }

  // 更新 theme variant 并且 re-render
  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }

  public render(): void {
    const semanticColors: Readonly<ISemanticColors> | undefined =
      this._themeVariant && this._themeVariant.semanticColors;
    const style: string = ` style="color:${semanticColors.bodyText}"`;
    
    this.domElement.innerHTML = `<p${'' || (this._themeProvider && style)}>Section background example.</p><p${'' || (this._themeProvider && style)}>Change the theme (Site Settings -> Change the look) and see what\'s happenning.</p>`;
  }
}
