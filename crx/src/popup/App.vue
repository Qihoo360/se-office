<template>
  <div class="popup">
    <div id="panel-createnew">
      <div class="header">新建</div>
      <div class="thumb-list">
        <div
          class="thumb-wrap"
          template="WORD"
          data-hint="2"
          data-hint-direction="left-top"
          data-hint-offset="22, 13"
          @click="onCreateNew('.docx')"
        >
          <div class="thumb" style="background-image: url(./img/doc-formats/docx.png)"></div>
          <div class="title" data-original-title="" title="">文档</div>
        </div>
        <div
          class="thumb-wrap"
          template="EXCEL"
          data-hint="2"
          data-hint-direction="left-top"
          data-hint-offset="22, 13"
          @click="onCreateNew('.xlsx')"
        >
          <div class="thumb" style="background-image: url(./img/doc-formats/xlsx.png)"></div>
          <div class="title" data-original-title="" title="">表格</div>
        </div>
        <div
          class="thumb-wrap"
          template="PPT"
          data-hint="2"
          data-hint-direction="left-top"
          data-hint-offset="22, 13"
          @click="onCreateNew('.pptx')"
        >
          <div class="thumb" style="background-image: url(./img/doc-formats/pptx.png)"></div>
          <div class="title" data-original-title="" title="">演示文稿</div>
        </div>
      </div>
      <div class="header">打开</div>
      <div class="open-container">
        <el-button
          type="info"
          size="large"
          :disabled="btnOpenDisabled"
          :icon="FolderOpened"
          @click="onOpenDocument"
          plain
        >
          打开本地文件
        </el-button>
      </div>
    </div>
  </div>
</template>
<script lang="ts" setup>
import { FolderOpened } from '@element-plus/icons-vue'
import { Extension } from '@/common/extension'
import { ref } from 'vue'
// import { Plugin } from '@/editor/plugin'

const btnOpenDisabled = ref(false)

const onCreateNew = (ext: string) => {
  chrome.tabs.create({ url: chrome.extension.getURL('app.html') + '?url=' + ext })
}

let timer: number
const onOpenDocument = async () => {
  btnOpenDisabled.value = true

  clearTimeout(timer)
  timer = window.setTimeout(() => {
    btnOpenDisabled.value = false
  }, 10 * 1000)

  const data = await Extension.sendToBackground('fileSelect', { type: '' })
  // const data = await Plugin.fileSelect('.docx')
  btnOpenDisabled.value = false

  const path = data?.data?.path
  if (path) {
    Extension.openUrl(path)
  }
}
</script>
<style lang="scss">
:root {
  --text-normal-pressed: fade(#000, 80%);

  --highlight-button-hover: #e0e0e0;
  --highlight-button-pressed: #cbcbcb;
  --highlight-button-pressed-hover: #bababa;
}
html,
body {
  margin: 0;
}
</style>
<style lang="scss" scoped>
$text-normal-pressed: var(--text-normal-pressed);

$highlight-button-hover: var(--highlight-button-hover);
$highlight-button-pressed: var(--highlight-button-pressed);
$highlight-button-pressed-hover: var(--highlight-button-pressed-hover);

#panel-createnew {
  width: 360px;

  .header {
    font-size: 18px;
    padding: 0 0 0 25px;
    white-space: nowrap;
    margin-top: 20px;
    margin-bottom: 20px;
  }

  .thumb-list {
    max-width: 600px;
    .thumb-wrap,
    .blank-document {
      display: inline-block;
      text-align: center;
      width: auto;
      cursor: pointer;
      vertical-align: top;
      // .border-radius(@border-radius-small);

      .thumb,
      .blank-document-btn {
        width: 96px;
        height: 96px;
        background-repeat: no-repeat;
        background-position: center;
        margin: 12px 12px 0px 12px;
        background-size: contain;
      }

      .title {
        width: 104px;
        // .font-size-medium();
        font-size: 14px;
        line-height: 14px;
        height: 28px;
        margin: 8px 8px 12px 8px;
        word-break: break-word;
        word-wrap: break-word;
        overflow: hidden;
        text-overflow: ellipsis;
        display: -webkit-box;
        -webkit-line-clamp: 2;
        -webkit-box-orient: vertical;
      }
      &:hover {
        background-color: $highlight-button-hover;
      }
      &:active {
        color: $text-normal-pressed;
        background-color: $highlight-button-pressed;
      }
    }
  }

  .open-container {
    text-align: center;
    padding-bottom: 25px;
  }
}
</style>
