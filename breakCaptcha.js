//FUNCAO PARA PEGAR INFORMACOES DE UM CAPTCHA COM GRID DE IMAGEM
async function catchGridCaptchaInfo(
    webdriver,
    captchaIframeXpath,
    saveCreenshotPath
  ) {
    try {
      // take a screenchot of the frame element
      const iframeScreenshot = await webdriver.takeScreenshot(
        "xpath",
        captchaIframeXpath
      );
      // create a buffer for the screenshot
      const iframeScreenshotBuffer = await sharp(
        Buffer.from(iframeScreenshot, "base64")
      ).toBuffer();
      // write a file
      fs.writeFileSync(
        `${saveCreenshotPath}/iframe-screenshot.png`,
        iframeScreenshotBuffer
      );
      //switch to frame
      await webdriver.changeFrame("xpath", captchaIframeXpath);
      // DECLARANDO VARIAVEIS COM ELEMENTOS DO CAPTCHA
      let captchaImagesRows = await webdriver.findElements(
        "xpath",
        `/html/body/div/div/div[2]/div[2]/div/table/tbody/tr[not(contains(@class, 'printonly'))]`
      );
      let captchaImagesColumns = await webdriver.findElements(
        "xpath",
        `//*[@id="rc-imageselect-target"]/table/tbody/tr[1]/td[not(contains(@class, 'printonly'))]`
      );
      let rowsCaptchaTable = captchaImagesRows.length;
      let columnsCaptchaTable = captchaImagesColumns.length;
      // find captcha header element
      let headerText = await webdriver.getText(
        "xpath",
        "/html/body/div/div/div[2]/div[1]/div[1]/div"
      );
      // EXTRAI TEXTO DE INSTRUÇÃO PARA RESOLVER O CAPTCHA
      let capchaHeaderElement;
      // VERIFICA SE AS INSTRUÇÕES FORAM EXTRAIDAS DO ELEMENTO ERRADO
      if (headerText == "") {
        // EXTRAI TEXTO DE INSTRUÇÃO PARA RESOLVER O CAPTCHA
        headerText = (
          await webdriver.getText(
            "xpath",
            "/html/body/div/div/div[2]/div[1]/div[1]/div"
          )
        ).replace(/\n/g, " ");
      }
      // VERIFICA SE AS INSTRUÇÕES FORAM EXTRAIDAS DO ELEMENTO ERRADO
      else if (headerText == "") {
        headerText = (
          await webdriver.getText(
            "xpath",
            `/html/body/div/div/div[2]/div[1]/div[1]/div[2]`
          )
        ).replace(/\n/g, " ");
      }
      // extract text from captcha instructions
      headerText = headerText.replace(/\n/g, " ");
      // get body element to crop the screenshot to show just the captcha image
      if (
        headerText.includes("nenhuma") ||
        headerText.includes("nenhum") ||
        headerText.length == 0
      ) {
        throw new Error(`Esse modelo de captcha não é possível tratar`);
      }
      const capchaBodyElement = await webdriver.findElement(
        "xpath",
        "/html/body/div/div/div[2]/div[2]/div"
      );
      // Get the location and size of the element
      const rectBody = await capchaBodyElement.getRect();
      const locationBody = { x: rectBody.x, y: rectBody.y };
      const sizeBody = { width: rectBody.width, height: rectBody.height };
      // Use sharp to crop the screenshot to include only the target element
      const captchaBodyScreenshotBuffer = await sharp(
        Buffer.from(iframeScreenshot, "base64")
      )
        .extract({
          left: locationBody.x,
          top: locationBody.y,
          width: sizeBody.width,
          height: sizeBody.height,
        })
        .toBuffer();
      // Save the cropped screenshot as an image file
      fs.writeFileSync(
        `${saveCreenshotPath}/captchaImage.png`,
        captchaBodyScreenshotBuffer
      );
      // RETORNA AS INSTRUÇÕES, O CAMINHO COMPLETO DO ARQUIVO COM A IMAGEM DO GRID DO CAPTCHA, O XPATH PARA O ELEMENTO DO CAPTCHA, A QUANTIDADE DE LINHAS E DE COLUNAS DO CAPTCHA
      return [
        headerText,
        `${saveCreenshotPath}/captchaImage.png`,
        "/html/body/div/div/div[2]/div[2]/div/table",
        rowsCaptchaTable,
        columnsCaptchaTable,
      ];
    } catch (err) {
      console.log(`Erro ao extrair informações do captcha${err}`);
      if (err.message.includes(`Esse modelo de captcha não`)) {
        throw new Error(`Esse modelo de captcha não é possível tratar`);
      }
      if (err.message.includes(`Error finding element`)) {
        console.log("Captcha não solicitado");
        throw new Error("Não tem captcha");
      }
      throw new Error(`Erro ao tirar print da imagem do captcha:\n${err}`);
    }
  }