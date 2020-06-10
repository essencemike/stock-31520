<template>
  <div class="wrapper">
    <span class="label">股票代码:</span>
    <input v-model="code" />
    <button class="gen-btn" @click="genTable">生成全景表</button>
    <p class="msg" v-if="loading">正在生成全景表...</p>
    <p class="msg" v-if="error">获取数据发生错误， 请联系管理员 email: gzxessence@163.com</p>
    <template v-if="stock">
      <h2 class="title">
        {{ stock.info.name }} ({{ stock.info.code }}) 全景表
        <span class="unit">单位: 亿元</span>
      </h2>
      <div class="time">日期: {{ stock.datetime }}</div>
      <div class="header-title">
        <span>{{ stock.company.GSname }}</span>
        <span>地域: {{ stock.company.GSdy }}</span>
        <span>行业: {{ stock.company.GShy }}</span>
        <span>上市日期: {{ stock.company.Sssj }}</span>
      </div>
      <div class="header-title">
        <span>主营: {{ stock.company.GSzy }}</span>
        <span>财报数据来源: 网易</span>
      </div>
      <table>
        <tbody>
          <tr v-for="(items, index) in stock.stock" :key="index" :class="{th: index === 0}">
            <td v-for="(value, i) in items" :key="i">{{ value }}</td>
          </tr>
        </tbody>
      </table>
      <div class="sep">
        <span>业绩(净利润)预期</span>
        <span>来源: 同花顺</span>
      </div>
      <div class="sep" v-if="!stock.yjyc || !stock.yjyc.length">本年度暂无机构做出业绩预测 </div>
      <table>
        <tbody>
          <tr v-for="(items, index) in stock.yjyc" :key="index" :class="{th: index === 0}">
            <td v-for="(value, i) in items" :key="i">{{ value }}</td>
          </tr>
        </tbody>
      </table>
    </template>

    <p class="tips">纯属娱乐，不构成任何投资建议，如有问题请联系我 email: gzxessence@163.com</p>
  </div>
</template>

<script>
import axios from 'axios';
import dayjs from 'dayjs';

import { errorCaptured } from '../utils';

import './index.less';

export default {
  name: 'Stock31520',
  data() {
    return {
      code: '',
      loading: false,
      stock: null,
      error: false,
    }
  },

  beforeDestory() {
    this.stock = null;
    this.loading = false;
    this.error = false;
  },

  methods: {
    async genTable() {
      this.loading = true;
      this.error = false;

      await this.setStockInfo(this.code);

      this.loading = false;
    },

    async setStockInfo(code) {
      const [err, data] = await errorCaptured(this.getStock(code));

      if (err) {
        console.error('获取数据发生错误: ', err);
        this.error = true;
        return;
      }

      // 初始化 stock
      this.stock = data;
      // 日期
      this.stock.datetime = dayjs().format('YYYY-MM-DD');
    },

    getStock(code) {
      const url = `/api/stock/${code}`;
      return new Promise((resolve, reject) => {
        axios(url).then(res => {
          const data = res.data;
          if (data.code === 200) {
            resolve(data.data);
          } else {
            reject(res)
          }
        }).catch(reject);
      });
    },
  },
}
</script>
