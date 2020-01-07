import RPA from "ts-rpa";
import { WebDriver, By, FileDetector, Key } from "selenium-webdriver";
import { DH_NOT_SUITABLE_GENERATOR } from "constants";
var fs = require('fs'); 


// 読み込みする スプレッドシートID と シート名 の記載
const SSID = process.env.Youtube_Card_Setting_SheetID;
const SSName1 = process.env.Youtube_Card_Setting_SheetName;


async function Start(){

    await RPA.Google.authorize({
        //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
        refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
        tokenType: "Bearer",
        expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
    });
    // 設定するアカウントデータ
    const Account1 = ["アカウント名","ID","PW","URL"];
    const Account2 = ["アカウント名","ID","PW","URL"];
    await AccountSelect(Account1,Account2);
    //そもそも設定するかどうかのフラッグ・URLか再生リストかのフラッグ・何秒設定・URL・フレーズ・ティーザー・検索タイトル/カードタイトル・設定タイトル/詳細・固定コメント
    const SheetDatas1 = [
    ["false",["URLor再生リスト"],["秒数"],["URL"],["フレーズ文言"],["ティーザー文言"],["検索タイトル・カードタイトル"],["動画タイトル/詳細"],["固定コメント"]],
    ["false",["URLor再生リスト"],["秒数"],["URL"],["フレーズ文言"],["ティーザー文言"],["検索タイトル・カードタイトル"],["動画タイトル/詳細"],["固定コメント"]],
    ["false",["URLor再生リスト"],["秒数"],["URL"],["フレーズ文言"],["ティーザー文言"],["検索タイトル・カードタイトル"],["動画タイトル/詳細"],["固定コメント"]]];
    const SheetDatas2 = [
    ["false",["URLor再生リスト"],["秒数"],["URL"],["フレーズ文言"],["ティーザー文言"],["検索タイトル・カードタイトル"],["動画タイトル/詳細"],["固定コメント"]],
    ["false",["URLor再生リスト"],["秒数"],["URL"],["フレーズ文言"],["ティーザー文言"],["検索タイトル・カードタイトル"],["動画タイトル/詳細"],["固定コメント"]],
    ["false",["URLor再生リスト"],["秒数"],["URL"],["フレーズ文言"],["ティーザー文言"],["検索タイトル・カードタイトル"],["動画タイトル/詳細"],["固定コメント"]]];
    // スプレッドシートからデータを取得する
    await ReadSheet1(SheetDatas1);
    await ReadSheet2(SheetDatas2);
    // 設定シート左側の開始フラグがtrueならログインして、設定処理する
    for(let i in SheetDatas1){
        if(SheetDatas1[i][0] == "true"){
            await YoutubeLogin(Account1);
            break
        }
    }
    for(let i in SheetDatas1){
        if(SheetDatas1[i][0] == "true"){
            await TitleSearch(SheetDatas1[i][6][0]);
            await RPA.sleep(3000);
            await CardCreate(SheetDatas1[i]);
            await TitleSyousaiChange(SheetDatas1[i]);
            if(SheetDatas1[i][8][0] != undefined){
                await CommentPost(SheetDatas1[i]);
            }
        }
    }
    // 設定シート右側の開始フラグがtrueならログインして、設定処理する
    for(let i in SheetDatas2){
        if(SheetDatas2[i][0] == "true"){
            await Youtube2ndLogin(Account2);
            break
        }
    }
    for(let i in SheetDatas2){
        if(SheetDatas2[i][0] == "true"){
            await TitleSearch(SheetDatas2[i][6][0]);
            await RPA.sleep(3000);
            await CardCreate(SheetDatas2[i]);
            await TitleSyousaiChange(SheetDatas2[i]);
            if(SheetDatas2[i][8][0] != undefined){
                await CommentPost(SheetDatas2[i]);
            }
        }
    }
}

Start()


async function AccountSelect(Account1,Account2){
    // 設定するアカウントを取得する
    const firstAccountData = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName1}!B2:S10`});
    // スプシB2・J2　のプルダウンをアカウントとして保持する
    Account1[0] = firstAccountData[0][0];
    Account2[0] = firstAccountData[0][8];
    for(let i in firstAccountData){
        // B2・J2　のプルダウンで選択されたアカウントのID・PW・URLを取得する
        if(Account1[0] == firstAccountData[i][14]){
            Account1[1] = firstAccountData[i][15];
            Account1[2] = firstAccountData[i][16];
            Account1[3] = firstAccountData[i][17];
        }
        if(Account2[0] == firstAccountData[i][14]){
            Account2[1] = firstAccountData[i][15];
            Account2[2] = firstAccountData[i][16];
            Account2[3] = firstAccountData[i][17];
        }
    }
    RPA.Logger.info("左アカウント "+Account1);
    RPA.Logger.info("右アカウント "+Account2);

}


async function YoutubeLogin(Account){
    await RPA.WebBrowser.get(Account[3]);
    await RPA.sleep(1000);
    //const GoogleLoginID = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementsLocated({className:'whsOnd zHQkBf'}),5000);
    const GoogleLoginID = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({id:"Email"}),8000);
    await RPA.WebBrowser.sendKeys(GoogleLoginID,[`${Account[1]}`]);
    const NextButton = await RPA.WebBrowser.findElementById('next');
    await RPA.WebBrowser.mouseClick(NextButton);
    await RPA.sleep(1000);
    const GoogleLoginPW = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({id:'Passwd'}),5000);
    await RPA.WebBrowser.sendKeys(GoogleLoginPW,[`${Account[2]}`]);
    const NextButton2 = await RPA.WebBrowser.findElementById('signIn');
    await RPA.WebBrowser.mouseClick(NextButton2);
    await RPA.sleep(4000);
}

async function Youtube2ndLogin(Account){
    // 別のアカウントに切り替えるためにログアウトする
    await RPA.WebBrowser.get("https://studio.youtube.com/channel/");
    await RPA.sleep(1500);
    const AccountButton = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({id:'avatar-btn'}),8000);
    await RPA.WebBrowser.mouseClick(AccountButton);
    await RPA.sleep(500);
    const LogOutButton = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementsLocated({className:'style-scope ytd-compact-link-renderer'}),5000);
    for(let i in LogOutButton){
        const LogOutText = await LogOutButton[i].getText();
        if(LogOutText == "ログアウト"){
            await RPA.WebBrowser.mouseClick(LogOutButton[i]);
            await RPA.sleep(3000);
            break
        }
    }
    // 別のアカウントでログインする
    const LoginButton = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementsLocated({className:'style-scope ytd-button-renderer style-suggestive size-small'}),5000);
    for(let i in LoginButton){
        const LoginButtonText = await LoginButton[i].getText();
        if(LoginButtonText == "ログイン"){
            await RPA.WebBrowser.mouseClick(LoginButton[i]);
            break
        }
    }
    await RPA.sleep(1500);
    const AccountAdd = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({id:"account-chooser-add-account"}),5000);
    await RPA.WebBrowser.mouseClick(AccountAdd);
    /*
    const AccountSelectButton = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementsLocated({className:"BHzsHc"}),5000);
    for(let i in AccountSelectButton){
        const Text1 = await AccountSelectButton[i].getText();
        if(Text1 == "別のアカウントを使用"){
            await RPA.WebBrowser.mouseClick(AccountSelectButton[i]);
            break
        }
    }
    */
    await RPA.sleep(1000);
    const LoginID = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({id:'Email'}),8000);
    await RPA.WebBrowser.sendKeys(LoginID,[Account[1]]);
    const NextButton1 = await RPA.WebBrowser.findElementById("next");
    await RPA.WebBrowser.mouseClick(NextButton1);
    await RPA.sleep(1000);
    const LoginPW = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({id:'Passwd'}),5000);
    await RPA.WebBrowser.sendKeys(LoginPW,[Account[2]]);
    const NextButton2 = await RPA.WebBrowser.findElementById("signIn");
    await RPA.WebBrowser.mouseClick(NextButton2);
    await RPA.sleep(4000);
    await RPA.WebBrowser.get(Account[3]);
    await RPA.sleep(3000);
}


async function TitleSearch(TargetTitle){
    const SearchInput = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({id:'query-input'}),9000);
    await RPA.sleep(1000);
    await RPA.WebBrowser.sendKeys(SearchInput,[TargetTitle]);
    await RPA.sleep(2000);
    const Titles = await RPA.WebBrowser.findElementsByClassName("style-scope ytcp-video-list-cell-video remove-default-style");
    const TitlesImage = await RPA.WebBrowser.findElementsByClassName("style-scope ytcp-img-with-fallback");
    for(let i in Titles){
        const TitleText = await Titles[i].getText();
        if(TitleText == TargetTitle){
            await RPA.WebBrowser.mouseClick(TitlesImage[i]);
            await RPA.sleep(3000);
            break
        }
    }
    await RPA.sleep(1000);
    const CardButton = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementsLocated({className:'dropdown-trigger-text style-scope ytcp-text-dropdown-trigger'}),5000);
    for(let i in CardButton){
        const CardButtonText = await CardButton[i].getText();
        if(CardButtonText.includes("カード") == true){
            await RPA.WebBrowser.mouseClick(CardButton[i]);
            await RPA.sleep(2000);
            break
        }
    }
    // 新しいタブができたらそちらに移る
    const windows = await RPA.WebBrowser.getAllWindowHandles()
    RPA.Logger.info(windows.length);
    if(windows.length != 1){
        await RPA.WebBrowser.switchToWindow(windows[1]);
    }

}


async function ReadSheet1(SheetDatas){
    const Data1 = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName1}!B3:F23`});
    const Data2 = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName1}!B26:F46`});
    const Data3 = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName1}!B49:F69`});
    const AllData = [];
    AllData[0] = Data1;
    AllData[1] = Data2;
    AllData[2] = Data3;
    for(let i in AllData){
        if(AllData[i][0][0] == "RPA実行OK"){
            SheetDatas[i][0] = "true";
            try{
                // 秒数が入っているブロックだけデータ取得する
                if(AllData[i][2][3].length > 0){
                    // リンクか再生リスト
                    SheetDatas[i][1][0] = AllData[i][1][0];
                    // 秒数
                    //SheetDatas[i][2][0] = AllData[i][2][3];
                    // URL
                    SheetDatas[i][3][0] = AllData[i][2][2];
                    // フレーズ文言
                    SheetDatas[i][4][0] = AllData[i][3][1];
                    // ティーザー文言
                    SheetDatas[i][5][0] = AllData[i][4][1];
                    // 検索するタイトル
                    SheetDatas[i][6][0] = AllData[i][2][4];
                    // カードのタイトル
                    SheetDatas[i][6][1] = AllData[i][2][1];
                    // カードのタイトルが 50文字以上なら、短くする
                    if(SheetDatas[i][6][1].length > 50){
                        RPA.Logger.info("タイトル文字数オーバー");
                        const splitdata = SheetDatas[i][6][1].split("|");
                        SheetDatas[i][6][1] = splitdata[0];
                    }
                    // 動画タイトル
                    SheetDatas[i][7][0] = AllData[i][4][4];
                    // 動画詳細
                    SheetDatas[i][7][1] = AllData[i][6][4];
                    // 固定コメント
                    SheetDatas[i][8][0] = AllData[i][8][4];
                }
                if(AllData[i][6][3].length > 0){
                    // リンクか再生リストか
                    SheetDatas[i][1][1] = AllData[i][5][0];
                    // 秒数
                    SheetDatas[i][2][0] = AllData[i][6][3];
                    // URL
                    SheetDatas[i][3][1] = AllData[i][6][2];
                    // フレーズ
                    SheetDatas[i][4][1] = AllData[i][7][1];
                    // ティーザー
                    SheetDatas[i][5][1] = AllData[i][8][1];
                    // カードタイトル
                    SheetDatas[i][6][1] = AllData[i][2][1];
                }
                if(AllData[i][10][3].length > 0){
                    SheetDatas[i][1][2] = AllData[i][9][0];
                    SheetDatas[i][2][1] = AllData[i][10][3];
                    SheetDatas[i][3][2] = AllData[i][10][2];
                    SheetDatas[i][4][2] = AllData[i][11][1];
                    SheetDatas[i][5][2] = AllData[i][12][1];
                    SheetDatas[i][6][1] = AllData[i][10][1];
                }
                if(AllData[i][14][3].length > 0){
                    SheetDatas[i][1][3] = AllData[i][13][0];
                    SheetDatas[i][2][2] = AllData[i][14][3];
                    SheetDatas[i][3][3] = AllData[i][14][2];
                    SheetDatas[i][4][3] = AllData[i][15][1];
                    SheetDatas[i][5][3] = AllData[i][16][1];
                    SheetDatas[i][6][1] = AllData[i][14][1];
                }
                if(AllData[i][18][3].length > 0){
                    SheetDatas[i][1][4] = AllData[i][17][0];
                    SheetDatas[i][2][3] = AllData[i][18][3];
                    SheetDatas[i][3][4] = AllData[i][18][2];
                    SheetDatas[i][4][4] = AllData[i][19][1];
                    SheetDatas[i][5][4] = AllData[i][20][1];
                    SheetDatas[i][6][1] = AllData[i][18][1];
                }
            }
            catch{
                ;
            }
        }
    }
    //RPA.Logger.info(SheetDatas);
}


async function ReadSheet2(SheetDatas){
    const Data1 = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName1}!J3:N23`});
    const Data2 = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName1}!J26:N46`});
    const Data3 = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName1}!J49:N69`});
    const AllData = [];
    AllData[0] = Data1;
    AllData[1] = Data2;
    AllData[2] = Data3;
    for(let i in AllData){
        if(AllData[i][0][0] == "RPA実行OK"){
            SheetDatas[i][0] = "true";
            try{
                // 秒数が入っているブロックだけデータ取得する
                if(AllData[i][2][3].length > 0){
                    // リンクか再生リスト
                    SheetDatas[i][1][0] = AllData[i][1][0];
                    // 秒数
                    //SheetDatas[i][2][0] = AllData[i][2][3];
                    // URL
                    SheetDatas[i][3][0] = AllData[i][2][2];
                    // フレーズ文言
                    SheetDatas[i][4][0] = AllData[i][3][1];
                    // ティーザー文言
                    SheetDatas[i][5][0] = AllData[i][4][1];
                    // 検索するタイトル
                    SheetDatas[i][6][0] = AllData[i][2][4];
                    // カードのタイトル
                    SheetDatas[i][6][1] = AllData[i][2][1];
                    // カードのタイトルが 50文字以上なら、短くする
                    if(SheetDatas[i][6][1].length > 50){
                        RPA.Logger.info("タイトル文字数オーバー");
                        const splitdata = SheetDatas[i][6][1].split("|");
                        SheetDatas[i][6][1] = splitdata[0];
                    }
                    // 動画タイトル
                    SheetDatas[i][7][0] = AllData[i][4][4];
                    // 動画詳細
                    SheetDatas[i][7][1] = AllData[i][6][4];
                    // 固定コメント
                    SheetDatas[i][8][0] = AllData[i][8][4];
                }
                if(AllData[i][6][3].length > 0){
                    // リンクか再生リストか
                    SheetDatas[i][1][1] = AllData[i][5][0];
                    // 秒数
                    SheetDatas[i][2][0] = AllData[i][6][3];
                    // URL
                    SheetDatas[i][3][1] = AllData[i][6][2];
                    // フレーズ
                    SheetDatas[i][4][1] = AllData[i][7][1];
                    // ティーザー
                    SheetDatas[i][5][1] = AllData[i][8][1];
                    // カードタイトル
                    SheetDatas[i][6][1] = AllData[i][2][1];
                }
                if(AllData[i][10][3].length > 0){
                    SheetDatas[i][1][2] = AllData[i][9][0];
                    SheetDatas[i][2][1] = AllData[i][10][3];
                    SheetDatas[i][3][2] = AllData[i][10][2];
                    SheetDatas[i][4][2] = AllData[i][11][1];
                    SheetDatas[i][5][2] = AllData[i][12][1];
                    SheetDatas[i][6][1] = AllData[i][10][1];
                }
                if(AllData[i][14][3].length > 0){
                    SheetDatas[i][1][3] = AllData[i][13][0];
                    SheetDatas[i][2][2] = AllData[i][14][3];
                    SheetDatas[i][3][3] = AllData[i][14][2];
                    SheetDatas[i][4][3] = AllData[i][15][1];
                    SheetDatas[i][5][3] = AllData[i][16][1];
                    SheetDatas[i][6][1] = AllData[i][14][1];
                }
                if(AllData[i][18][3].length > 0){
                    SheetDatas[i][1][4] = AllData[i][17][0];
                    SheetDatas[i][2][3] = AllData[i][18][3];
                    SheetDatas[i][3][4] = AllData[i][18][2];
                    SheetDatas[i][4][4] = AllData[i][19][1];
                    SheetDatas[i][5][4] = AllData[i][20][1];
                    SheetDatas[i][6][1] = AllData[i][18][1];
                }
            }
            catch{
                ;
            }
        }
    }
    //RPA.Logger.info(SheetDatas);
}


async function CardCreate(CardData){
    await RPA.sleep(200);
    const MoviePlayer = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({id:'movie_player'}),5000);
    const MoviePlayButton = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({className:'ytp-play-button ytp-button'}),3000);
    await RPA.WebBrowser.mouseMove(MoviePlayButton);
    // 開始3秒で一度止めてカードを設定する
    while(0 == 0){
        const TimeSecond_3 = await RPA.WebBrowser.findElementByClassName("ytp-time-current");
        const TimeText = await TimeSecond_3.getText();
        if(String(TimeText) == '0:03'){
            await RPA.WebBrowser.mouseClick(MoviePlayer);
            break
        }
        await RPA.sleep(200);
    }
    RPA.Logger.info(CardData);
    // カードを追加するボタンをおす
    const CardAppendButton = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:"/html/body/div[2]/div[4]/div/div[5]/div/div/div[1]/div[2]/div[2]/div[1]/span/button"}),5000);
    const CardAppendButtonText = await CardAppendButton.getText();
    if(String(CardAppendButtonText) == "カードを追加"){
        await RPA.WebBrowser.mouseClick(CardAppendButton);
        await RPA.sleep(300);
    }
    // リンク作成 ボタンをおす
    const CardCreateButton = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementsLocated({className:'yt-uix-button yt-uix-button-size-default yt-uix-button-default annotator-create-button  yt-uix-clickcard-close'}),5000);
    const CardCreateButtonText = await CardCreateButton[3].getText();
    if(CardCreateButtonText == "作成"){
        await RPA.WebBrowser.mouseClick(CardCreateButton[3]);
        await RPA.sleep(300);
    }
    // リンクのURL　の文字が出るまで待機
    while(0 == 0){
        try{
            const LinkElement = await RPA.WebBrowser.findElementByClassName("annotator-overlay-input-label");
            const LinkURLText = await LinkElement.getText();
            if(LinkURLText == "リンクの URL"){
                break
            }
        }
        catch{
            ;
        }
    }
    await RPA.sleep(1000);
    // リンクURLに入力する
    const LinkURLInput = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({className:'yt-uix-form-input-text annotator-prs-url annotator-initial-focus annotator-overlay-text-input'}),5000);
    await RPA.WebBrowser.sendKeys(LinkURLInput,[CardData[3][0]]);
    await RPA.sleep(300);
    const CardNextButton = await RPA.WebBrowser.findElementByClassName("yt-uix-button yt-uix-button-size-default yt-uix-button-primary annotator-next");
    await RPA.WebBrowser.mouseClick(CardNextButton);
    // 関連ウェブサイト　の文字が出てくるまで待機
    while(0 == 0){
        await RPA.sleep(500);
        try{
            const Title = await RPA.WebBrowser.findElementByClassName("yt annotator-prs-card-details create-only");
            const TitleText = await Title.getText();
            if(TitleText == "関連ウェブサイト"){
                RPA.Logger.info("関連ウェブサイトの文字ありました");
                break
            }
        }
        catch{
            ;
        }
    }
    // スプレッドシートにカードのタイトル記載があれば、入力する
    const CardTitleInput = await RPA.WebBrowser.findElementById("annotator-prs-title-associated");
    try{
        if(String(CardData[6][1]).length > 1){
            RPA.Logger.info("タイトル記載されています")
            await CardTitleInput.clear();
            await RPA.sleep(200);
            await RPA.WebBrowser.sendKeys(CardTitleInput,[CardData[6][1]]);
            await RPA.sleep(300);
        }
        // スプレッドシートにカードタイトルの記載がなければ自動で入力された文字数を判定する
        if(String(CardData[6][1]).length < 1){
            RPA.Logger.info("タイトル記載されていません")
            const Text = await document.getElementById("annotator-prs-title-associated").getAttribute("value");
            await RPA.sleep(1000);
            RPA.Logger.info(Text);
            await RPA.WebBrowser.takeScreenshot();
            if(String(Text).length > 50){
                RPA.Logger.info("カードタイトル50文字超えたので調整します");
                const SplitText = await Text.split("|");
                await CardTitleInput.clear();
                await RPA.sleep(200);
                RPA.Logger.info(SplitText);
                await RPA.WebBrowser.sendKeys(CardTitleInput,[SplitText[0]]);
            }
        }
    }
    catch{
        RPA.Logger.info("3秒のカードタイトルがおかしいです");
    }
    // フレーズを入力する
    const PhraseInput = await RPA.WebBrowser.findElementById("annotator-prs-action-associated");
    if(CardData[4][0] != undefined){
        RPA.Logger.info("3秒のフレーズ入力します");
        await PhraseInput.clear();
        await RPA.sleep(100);
        await RPA.WebBrowser.sendKeys(PhraseInput,[CardData[4][0]]);
    }
    // ティーザーを入力する
    const TeaserInput = await RPA.WebBrowser.findElementById("annotator-teaser-text-associated");
    if(CardData[5][0] != undefined){
        RPA.Logger.info("3秒のティーザー入力します");
        await TeaserInput.clear();
        await RPA.sleep(100);
        await RPA.WebBrowser.sendKeys(TeaserInput,[CardData[5][0]]);
    }
    await RPA.WebBrowser.takeScreenshot();
    // カードを作成のボタンをおす
    const CardCreateApplyButton = await RPA.WebBrowser.findElementsByClassName("yt-uix-button yt-uix-button-size-default yt-uix-button-primary annotator-save create-only");
    for(let i in CardCreateApplyButton){
        RPA.Logger.info(await CardCreateApplyButton[i].getText());
        if(String(await CardCreateApplyButton[i].getText()) == "カードを作成"){
            await RPA.WebBrowser.mouseClick(CardCreateApplyButton[i]);
            await RPA.sleep(5000);
        }
    }
    await RPA.WebBrowser.sendKeys(MoviePlayer,[RPA.WebBrowser.Key.chord("j")]);

    // リンク設定が2つ以上あるか判定する
    const CreatesFlag = [];
    CreatesFlag[0] = CardData[2].length;
    if(CreatesFlag[0] == 0){
        RPA.Logger.info('3秒以外ないのでリターン');
        return
    }
    // 2つ目以降は、指定した秒数で設定する
    const OldTime = [];
    OldTime[0] = 0;
    for(let i in CardData[2]){
        const Timedata = String(CardData[2][i]).split(":");
        const minutedata = Number(Timedata[0])*6;
        const seconddata = Number(Timedata[1])/10;
        RPA.Logger.info(minutedata + seconddata + "0秒に設定");
        const SettingTime = minutedata + seconddata - OldTime[0];
        for(let v = 0; v < SettingTime;v++){
            await RPA.WebBrowser.sendKeys(MoviePlayer,[RPA.WebBrowser.Key.chord("l")]);
            await RPA.sleep(150);
        }
        await CardCreates(CardData,Number(i)+1);
        await RPA.sleep(5000);
        OldTime[0] = minutedata + seconddata;
    }
}


async function CardCreates(CardData,i){
    // カードを追加するボタンをおす
    const CardAppendButton = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:"/html/body/div[2]/div[4]/div/div[5]/div/div/div[1]/div[2]/div[2]/div[1]/span/button"}),5000);
    const CardAppendButtonText = await CardAppendButton.getText();
    if(String(CardAppendButtonText) == "カードを追加"){
        await RPA.WebBrowser.mouseClick(CardAppendButton);
        await RPA.sleep(300);
    }
    const Link_or_ListFlag = [];
    Link_or_ListFlag[0] = "None";
    // リンク作成か再生リスト作成か判定して処理を変える
    if(CardData[1][i] == "リンク"){
        Link_or_ListFlag[0] = "リンク";
    }
    if(CardData[1][i] == "再生リスト"){
        Link_or_ListFlag[0] = "再生リスト";
    }
    const CardCreateButton = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementsLocated({className:'yt-uix-button yt-uix-button-size-default yt-uix-button-default annotator-create-button  yt-uix-clickcard-close'}),5000);
    // フラグが再生リストなら、再生リスト作成を行う
    if(Link_or_ListFlag[0] == "再生リスト"){
        const CardCreateButtonText = await CardCreateButton[0].getText();
        if(CardCreateButtonText == "作成"){
            await RPA.WebBrowser.mouseClick(CardCreateButton[0]);
            await RPA.sleep(300);
        }
        const PlayListInput = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({className:'yt-uix-form-input-text yt-video-picker-url'}),5000);
        await RPA.WebBrowser.sendKeys(PlayListInput,[CardData[3][i]]);
        const TextCreateButton = await RPA.WebBrowser.findElementByXPath('//*[@id="annotations-edit-overlay-content-video"]/div/form/div[1]/details/summary');
        await RPA.WebBrowser.mouseClick(TextCreateButton);
        const ListPhraseInput = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({id:'annotator-custom-text-video'}),5000);
        await RPA.WebBrowser.sendKeys(ListPhraseInput,[CardData[4][i]]);
        const ListTeaserInput = await RPA.WebBrowser.findElementById("annotator-teaser-text-video");
        await RPA.WebBrowser.sendKeys(ListTeaserInput,[CardData[5][i]]);
        await RPA.sleep(300);
        const CardCreateApplyButton = await RPA.WebBrowser.findElementsByClassName("yt-uix-button yt-uix-button-size-default yt-uix-button-primary annotator-save create-only");
        for(let i in CardCreateApplyButton){
            const Text = await CardCreateApplyButton[i].getText();
            if(String(Text) == "カードを作成"){
                // カード作成のボタンをおす
                await RPA.WebBrowser.mouseClick(CardCreateApplyButton[i]);
                await RPA.sleep(3000);
            }
        }

    }
    // フラッグがリンクなら、リンク作成を行う
    if(Link_or_ListFlag[0] == "リンク"){
        const CardCreateButtonText = await CardCreateButton[3].getText();
        if(CardCreateButtonText == "作成"){
            await RPA.WebBrowser.mouseClick(CardCreateButton[3]);
            await RPA.sleep(300);
        }
        // リンクのURL　の文字が出るまで待機
        while(0 == 0){
            try{
                const LinkElement = await RPA.WebBrowser.findElementByClassName("annotator-overlay-input-label");
                const LinkURLText = await LinkElement.getText();
                if(LinkURLText == "リンクの URL"){
                    break
                }
            }
            catch{
                ;
            }
        }
        await RPA.sleep(1000);
        // リンクURLに入力する
        const LinkURLInput = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({id:'annotator-prs-url-link'}),5000);
        await RPA.WebBrowser.sendKeys(LinkURLInput,[CardData[3][i]]);
        await RPA.sleep(300);
        const CardNextButton = await RPA.WebBrowser.findElementByClassName("yt-uix-button yt-uix-button-size-default yt-uix-button-primary annotator-next");
        await RPA.WebBrowser.mouseClick(CardNextButton);
        // 関連ウェブサイト　の文字が出てくるまで待機
        while(0 == 0){
            await RPA.sleep(500);
            try{
                const Title = await RPA.WebBrowser.findElementByClassName("yt annotator-prs-card-details create-only");
                const TitleText = await Title.getText();
                if(TitleText == "関連ウェブサイト"){
                    RPA.Logger.info("関連ウェブサイトの文字ありました");
                    break
                }
            }
            catch{
                ;
            }
        }
        await RPA.sleep(500);
        // スプレッドシートにカードのタイトル記載があれば、入力する
        const CardTitleInput = await RPA.WebBrowser.findElementById("annotator-prs-title-associated");
        try{
            if(String(CardData[6][i]).length > 1){
                RPA.Logger.info("タイトル記載されています")
                await CardTitleInput.clear();
                await RPA.sleep(200);
                await RPA.WebBrowser.sendKeys(CardTitleInput,[CardData[6][i]]);
                await RPA.sleep(300);
            }
            // スプレッドシートにカードタイトルの記載がなければ自動で入力された文字数を判定する
            if(String(CardData[6][i]).length < 1){
                RPA.Logger.info("タイトル記載されていません")
                const Text = await (await RPA.WebBrowser.findElementById("annotator-prs-title-associated")).getAttribute("value");
                await RPA.sleep(1000);
                RPA.Logger.info(Text);
                await RPA.WebBrowser.takeScreenshot();
                if(String(Text).length > 50){
                    RPA.Logger.info("カードタイトル50文字超えたので調整します");
                    const SplitText = await Text.split("|");
                    await CardTitleInput.clear();
                    await RPA.sleep(200);
                    RPA.Logger.info(SplitText);
                    await RPA.WebBrowser.sendKeys(CardTitleInput,[SplitText[0]]);
                }
            }
        }
        catch{
            RPA.Logger.info("カードタイトルがおかしいです");
        }
        // フレーズを入力する
        const PhraseInput = await RPA.WebBrowser.findElementById("annotator-prs-action-associated");
        if(String(CardData[4][i]).length > 1){
            await PhraseInput.clear();
            await RPA.sleep(100);
            await RPA.WebBrowser.sendKeys(PhraseInput,[CardData[4][i]]);
        }
        // ティーザーを入力する
        const TeaserInput = await RPA.WebBrowser.findElementById("annotator-teaser-text-associated");
        if(String(CardData[5][i]).length > 1){
            await TeaserInput.clear();
            await RPA.sleep(100);
            await RPA.WebBrowser.sendKeys(TeaserInput,[CardData[5][i]]);
        }
        // カードを作成のボタンをおす
        const CardCreateApplyButton = await RPA.WebBrowser.findElementsByClassName("yt-uix-button yt-uix-button-size-default yt-uix-button-primary annotator-save create-only");
        for(let v in CardCreateApplyButton){
            RPA.Logger.info(await CardCreateApplyButton[v].getText());
            if(String(await CardCreateApplyButton[v].getText()) == "カードを作成"){
                await RPA.WebBrowser.mouseClick(CardCreateApplyButton[v]);
                await RPA.sleep(4000);
            }
        }
    }
}

// 動画のタイトルと詳細を記入する
async function TitleSyousaiChange(CardData){
    // タブが２つ以上あれば1以外消す
    const TabClose = await RPA.WebBrowser.getAllWindowHandles()
    RPA.Logger.info(TabClose.length);
    if(TabClose.length > 1){
        RPA.Logger.info("タブ閉じました")
        await RPA.WebBrowser.closeWindow(TabClose[1]);
        await RPA.WebBrowser.switchToWindow(TabClose[0]);
        await RPA.sleep(500);
    }
    const titleinput = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/ytcp-app/ytcp-entity-page/div/div/main/div/ytcp-animatable[8]/ytcp-video-metadata-editor-section/ytcp-video-metadata-editor/div/ytcp-animatable/ytcp-video-metadata-basics/div/div[1]/div[1]/ytcp-mention-textbox/ytcp-form-input-container/div[1]/div[2]/ytcp-mention-input/div'}),8000);
    const syousaiinput = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementsLocated({className:'style-scope ytcp-mention-textbox'}),5000);
    // 動画タイトルを一度エンコードして jsで入力する
    if(CardData[7][0] != undefined){
        await titleinput.clear();
        await RPA.sleep(300);
        const Text1 = await encodeURI(CardData[7][0]);
        await RPA.WebBrowser.driver.executeScript(`document.querySelectorAll("#textbox")[0].innerText = decodeURI("${Text1}")`);
        await RPA.WebBrowser.sendKeys(titleinput,[` `]);
        await RPA.Logger.info("タイトル入力したよ")
        await syousaiinput[9].clear();
        await RPA.sleep(300);
    }
    // 詳細を一度エンコードして jsで入力する
    if(CardData[7][1] != undefined){
        const Text2 = await encodeURI(CardData[7][1]);
        await RPA.WebBrowser.driver.executeScript(`document.getElementsByClassName("style-scope ytcp-mention-textbox")[9].innerText = decodeURI("${Text2}")`);
        await RPA.WebBrowser.sendKeys(syousaiinput[9],[` `]);
        await RPA.Logger.info("詳細入力したよ")
    }
    const SaveButton = await RPA.WebBrowser.findElementById("save");
    // 保存ボタンをおす
    await RPA.WebBrowser.mouseClick(SaveButton);
    await RPA.sleep(2300);
}

// 公式Youtubeページに遷移し、固定用コメントを投稿する
async function CommentPost(CardData){
    //await RPA.WebBrowser.get("https://studio.youtube.com/video/jEmlZ0qZaqI/edit?utm_campaign=upgrade&utm_medium=redirect&utm_source=%2Fmy_videos");
    //await RPA.sleep(2000);
    const YoutubeLink = await RPA.WebBrowser.findElementById("overlay-link-to-youtube");
    const YoutubeURL = await YoutubeLink.getAttribute("href");
    await RPA.WebBrowser.driver.executeScript("window.open(arguments[0], 'newtab')");
    await RPA.sleep(500);
    // 新しいタブを開き、そこでコメント投稿する
    const Tabs = await RPA.WebBrowser.getAllWindowHandles()
    if(Tabs.length > 1){
        await RPA.WebBrowser.switchToWindow(Tabs[1]);
        await RPA.sleep(200);
    }
    await RPA.WebBrowser.get(YoutubeURL);
    await RPA.sleep(2000);
    await RPA.WebBrowser.scrollTo({xpath:(`//*[@id="description"]/yt-formatted-string`)});
    await RPA.sleep(1000);
    const CommentInputButton = await RPA.WebBrowser.findElementById("placeholder-area");
    await RPA.WebBrowser.mouseClick(CommentInputButton);
    await RPA.sleep(500);
    const CommentInput = await RPA.WebBrowser.findElementById("contenteditable-root");
    // コメントを一度エンコードして jsで入力する
    const Text = await encodeURI(CardData[8][0]);
    await RPA.WebBrowser.driver.executeScript(`document.getElementById("contenteditable-root").innerText = decodeURI("${Text}")`);
    await RPA.sleep(100);
    // コメントの最後に半角スペースを入れると投稿できる
    await RPA.WebBrowser.sendKeys(CommentInput,[" "]);
    await RPA.sleep(200);
    const CommentButton = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementsLocated({className:'style-scope ytd-button-renderer style-primary size-default'}),5000);
    for(let i in CommentButton){
        const CommentButtonText = await CommentButton[i].getText();
        if(CommentButtonText == "コメント"){
            await RPA.WebBrowser.mouseClick(CommentButton[i]);
            break
        }
    }
    await RPA.sleep(2000);
    const CommentsCount = await RPA.WebBrowser.findElementsByCSSSelector("#action-buttons");
    RPA.Logger.info("コメント数 → "+CommentsCount.length);
    // 一番上のコメントを固定する
    const Comment1 = await RPA.WebBrowser.findElementByXPath("/html/body/ytd-app/div/ytd-page-manager/ytd-watch-flexy/div[4]/div[1]/div/ytd-comments/ytd-item-section-renderer/div[3]/ytd-comment-thread-renderer/ytd-comment-renderer/div[2]/div[3]/ytd-menu-renderer/yt-icon-button/button");
    await RPA.WebBrowser.mouseClick(Comment1);
    const KoteiButton = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementsLocated({className:'style-scope ytd-menu-navigation-item-renderer'}),5000);
    for(let i in KoteiButton){
        const KoteiText = await KoteiButton[i].getText();
        if(KoteiText == "固定"){
            await RPA.WebBrowser.mouseClick(KoteiButton[i]);
            break
        }
    }
    // 固定するボタンをおす
    const KoteiApplyButton = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementsLocated({className:'style-scope yt-button-renderer style-primary size-default'}),5000);
    await RPA.WebBrowser.mouseClick(KoteiApplyButton[0]);
    await RPA.sleep(2000);
    // タブが２つ以上あれば1以外消す
    const TabClose = await RPA.WebBrowser.getAllWindowHandles()
    RPA.Logger.info(TabClose.length);
    if(TabClose.length > 1){
        await RPA.WebBrowser.closeWindow(TabClose[1]);
        await RPA.WebBrowser.switchToWindow(TabClose[0]);
        await RPA.sleep(500);
    }

}