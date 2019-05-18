<template>
    <div class="theme" :style="bgProps">
        <div class="btn">
            <input type="color" v-model="changeColor">
        </div>
    </div>
</template>
<script>
export default {
    data() {
        return {
            changeColor: '#00cc00'
        }
    },
    computed: {
        bgProps() {
           return {
               '--bg-color': this.changeColor
           }
        }
    },
}
</script>
<style lang="scss" scoped>
.theme{
    width: 100%;
    height: 500px;
    display: flex;
    align-items: center;
    justify-content: center;
    background-color: var(--bg-color);
    .btn > input{
        width: 30px;
        height: 30px;
    }
}
</style>

